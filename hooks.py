import base64
import glob
import logging
import os
import time
from datetime import timedelta

import requests
from openpyxl import load_workbook
from psycopg2.errors import SerializationFailure

from odoo.fields import Command

_logger = logging.getLogger(__name__)

# New xml_ids created by this import (e.g. for a journal/asset row that has
# no existing match) are registered under this addon's own name, regardless
# of which client module's data actually triggered the import.
MODULE = os.path.basename(os.path.dirname(__file__))

JOURNAL_TYPES = {"Bank": "bank", "Kas": "cash", "Lain-lain": "general"}
ASSET_METHODS = {"Garis Lurus": "linear"}
ASSET_PERIODS = {"Bulan": "1", "Tahun": "12"}

# Used when the "company" sheet's external_report_layout row is blank -
# every client so far uses the same layout, but the sheet can still override it.
DEFAULT_EXTERNAL_REPORT_LAYOUT = "web.external_layout_folder"

# "company" sheet key -> res.config.settings field toggled through it. Both
# are applied via res.config.settings.execute(), not company.write(), since
# that's what actually applies the group implication / triggers the module
# install - same as checking the box in Settings and clicking Save.
ACCOUNTING_FEATURE_FIELDS = {
    "analytic_accounting": "group_analytic_accounting",
    "budget_management": "module_account_budget",  # installs account_budget
}

_TRUE_VALUES = {"true", "ya", "yes", "1", "aktif"}
_FALSE_VALUES = {"false", "tidak", "no", "0", "nonaktif", "non-aktif"}


def _parse_bool(value):
    """Parse a yes/no "company" sheet cell. None/blank means "not specified,
    don't touch" - unlike the other company fields, a boolean's off state is
    a real value someone might want to set, so this needs three outcomes
    (True/False/None) instead of just skip-if-blank.
    """
    if isinstance(value, bool):
        return value
    if value in (None, ""):
        return None
    text = str(value).strip().lower()
    if text in _TRUE_VALUES:
        return True
    if text in _FALSE_VALUES:
        return False
    _logger.warning('Unrecognized yes/no value %r in the "company" sheet - ignoring.', value)
    return None

# Set to True to auto-validate imported assets (generates the depreciation
# board and posts its moves). False leaves them as draft for manual review.
VALIDATE_IMPORTED_ASSETS = False

# Set to True to auto-post imported opening vendor/customer moves.
# False leaves them as draft for manual review.
POST_OPENING_MOVES = False


def find_data_file(data_dir):
    """Path to the single .xlsx under data_dir, or None if there isn't one -
    the data import is optional seed data, not a structural part of a client
    module, so a missing file just skips the import instead of blocking
    install. More than one file is treated as a real misconfiguration
    (which one would we import?) and still raises.
    """
    matches = glob.glob(os.path.join(data_dir, "*.xlsx"))
    if len(matches) > 1:
        raise ValueError(f"Expected at most one .xlsx file in {data_dir}, found {len(matches)}")
    return matches[0] if matches else None


def import_bundled_data(env, module_dir):
    """Convenience entry point for a client module's post_init_hook: find
    the single .xlsx under <module_dir>/data and run the import, or skip
    with a warning if there isn't one.
    """
    data_dir = os.path.join(module_dir, "data")
    path = find_data_file(data_dir)
    if not path:
        _logger.warning("No .xlsx master data file found under %s - skipping data import.", data_dir)
        return
    wb = load_workbook(path, read_only=True, data_only=True)
    run_import(env, wb)


def _read_opening_balance_date(wb):
    """Read OPENING_BALANCE_DATE, a key/value row in the "company" sheet
    (falls back to the older "petunjuk" location for workbooks that haven't
    been migrated yet). Ini tanggal cutover opening balance yang sebenarnya
    (mis. "saldo per 30 Juni 2026") - fiscal_year_start di run_import()
    di-derive dari sini (+1 hari), bukan sebaliknya, karena inilah nilai
    yang paling natural diisi orang: tanggal per kapan saldo awal berlaku.
    """
    for sheet in ("company", "petunjuk"):
        if sheet not in wb.sheetnames:
            continue
        for row in wb[sheet].iter_rows(values_only=True):
            if row and row[0] == "OPENING_BALANCE_DATE":
                return row[1].date()
    raise ValueError('"OPENING_BALANCE_DATE" tidak ditemukan di sheet "company" atau "petunjuk"')


def _sheet_rows(wb, sheet, columns):
    rows = wb[sheet].iter_rows(values_only=True)
    header = next(rows)
    # A column entirely absent from the header (not just blank cells) reads
    # as None for every row - same tolerance _move_rows already gives an
    # optional column, so a sheet can omit one it doesn't need at all.
    idx = {col: header.index(col) for col in columns if col in header}
    for row in rows:
        values = {col: (row[idx[col]] if col in idx and idx[col] < len(row) else None) for col in columns}
        if values["id"]:
            values["id"] = values["id"] if "." in values["id"] else f"{MODULE}.{values['id']}"
            yield values


def _get_or_create(env, model, xml_id, values):
    record = env.ref(xml_id, raise_if_not_found=False)
    if record:
        record.write(values)
        return record
    # env.ref() returns None both when the xml_id was never registered and
    # when its target record was deleted without cleaning up ir.model.data
    # (e.g. manual deletion in the UI) - reuse/repoint a stale entry instead
    # of blindly creating a duplicate (module, name) row.
    record = env[model].create(values)
    module, name = xml_id.split(".", 1)
    data = env["ir.model.data"].search([("module", "=", module), ("name", "=", name)])
    if data:
        data.write({"model": model, "res_id": record.id})
    else:
        env["ir.model.data"].create({"module": module, "name": name, "model": model, "res_id": record.id})
    return record


def _account_by_code_name(env, code_name):
    if not code_name:
        return env["account.account"]
    code = code_name.split(" ", 1)[0]
    return env["account.account"].search([("code", "=", code)], limit=1)


def _elapsed_depreciation_periods(acquisition_date, balance_date, method_period):
    """Whole periods (inclusive of the acquisition period) from
    acquisition_date through balance_date, for straight-line proration.
    """
    if acquisition_date > balance_date:
        return 0
    if method_period == "1":  # monthly
        return (balance_date.year - acquisition_date.year) * 12 + (balance_date.month - acquisition_date.month) + 1
    return balance_date.year - acquisition_date.year + 1  # yearly


def _write_opening_balances(env, totals):
    """totals: dict of account.account recordset -> [opening_debit, opening_credit].
    Writing these fields queues a rebuild of the company's single opening
    balance move, which Odoo refuses once that move is posted - skip
    entirely once the user has reviewed and posted it, instead of crashing.
    """
    opening_move = env.ref("base.main_company").account_opening_move_id
    if opening_move and opening_move.state != "draft":
        return
    for account, (debit, credit) in totals.items():
        values = {}
        if debit:
            values["opening_debit"] = debit
        if credit:
            values["opening_credit"] = credit
        if values:
            account.write(values)


def _partner_by_name(env, name, rank_field):
    if not name:
        return env["res.partner"]
    partner = env["res.partner"].search([("name", "=", name)], limit=1)
    return partner or env["res.partner"].create({"name": name, rank_field: 1})


def _move_rows(wb, sheet, default_date):
    """Group an Odoo-import-style sheet (a header row per move, followed by
    blank-id continuation rows) into one dict per move with a "lines" list.
    """
    rows = wb[sheet].iter_rows(values_only=True)
    header = next(rows)
    idx = {col: header.index(col) for col in header}

    def get(row, col):
        i = idx.get(col)
        return row[i] if i is not None and i < len(row) else None

    move = None
    for row in rows:
        xml_id = get(row, "id")
        if xml_id:
            if move:
                yield move
            xml_id = xml_id if "." in xml_id else f"{MODULE}.{xml_id}"
            raw_date = get(row, "date")
            move = {
                "id": xml_id,
                "ref": get(row, "Reference"),
                "date": raw_date.date() if raw_date else default_date,
                "journal": get(row, "journal"),
                "lines": [],
            }
        if move is None:
            continue
        account = get(row, "line_ids/account")
        if account:
            move["lines"].append({
                "account": account,
                "debit": get(row, "line_ids/debit"),
                "credit": get(row, "line_ids/credit"),
                "name": get(row, "line_ids/name"),
                "partner": get(row, "line_ids/partner"),
                "date_maturity": get(row, "line_ids/date_maturity"),
            })
    if move:
        yield move


def _import_opening_move(env, wb, sheet, partner_rank_field, default_date):
    company = env.ref("base.main_company")
    for move_data in _move_rows(wb, sheet, default_date):
        existing = env.ref(move_data["id"], raise_if_not_found=False)
        if existing and existing.state != "draft":
            continue

        line_commands = []
        balance = 0.0
        for line in move_data["lines"]:
            debit, credit = line["debit"] or 0.0, line["credit"] or 0.0
            balance += credit - debit
            vals = {
                "account_id": _account_by_code_name(env, line["account"]).id,
                "debit": debit,
                "credit": credit,
                "name": line["name"],
                "partner_id": _partner_by_name(env, line["partner"], partner_rank_field).id,
            }
            if line["date_maturity"]:
                vals["date_maturity"] = line["date_maturity"].date()
            line_commands.append(Command.create(vals))

        if not company.currency_id.is_zero(balance):
            balancing_account = company.get_unaffected_earnings_account()
            line_commands.append(Command.create({
                "account_id": balancing_account.id,
                "debit": max(balance, 0.0),
                "credit": max(-balance, 0.0),
                "name": "Automatic Balancing Line",
            }))

        values = {
            "ref": move_data["ref"],
            "date": move_data["date"],
            "journal_id": env["account.journal"].search([("name", "=", move_data["journal"])], limit=1).id,
            "move_type": "entry",
            "line_ids": [Command.clear()] + line_commands,
        }
        move = _get_or_create(env, "account.move", move_data["id"], values)
        if POST_OPENING_MOVES and move.state == "draft":
            move.action_post()


def _import_company(env, wb):
    """Read the "company" sheet - key/value rows (also where
    OPENING_BALANCE_DATE lives, see _read_opening_balance_date) - and write
    it onto the single company record, then apply any accounting feature
    toggles found there (see
    ACCOUNTING_FEATURE_FIELDS). Optional sheet, and a blank cell leaves that
    field/toggle untouched rather than clearing it - so a partially-filled
    sheet (e.g. no logo URL yet) never blanks out data set another way.
    """
    if "company" not in wb.sheetnames:
        return
    data = {}
    for row in wb["company"].iter_rows(values_only=True):
        if row and row[0]:
            data[row[0]] = row[1]
    if not data:
        return

    company = env.ref("base.main_company")
    values = {}
    for field in ("name", "street", "street2", "city", "zip", "phone", "email", "website", "report_footer"):
        value = data.get(field)
        if value in (None, ""):
            continue
        # A numeric-looking text field (e.g. zip) can come back as a float
        # if the sheet cell got auto-formatted as a number - avoid writing
        # "55151.0" instead of "55151".
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        values[field] = str(value)

    if data.get("country"):
        country = env["res.country"].search([("name", "=", data["country"])], limit=1)
        values["country_id"] = country.id
        if data.get("state"):
            state = env["res.country.state"].search(
                [("name", "=", data["state"]), ("country_id", "=", country.id)], limit=1
            )
            values["state_id"] = state.id

    layout = env.ref(data.get("external_report_layout") or DEFAULT_EXTERNAL_REPORT_LAYOUT, raise_if_not_found=False)
    if layout:
        values["external_report_layout_id"] = layout.id

    logo_url = data.get("logo")
    if logo_url:
        try:
            response = requests.get(logo_url, timeout=10)
            response.raise_for_status()
            values["logo"] = base64.b64encode(response.content)
        except Exception:
            _logger.warning("Failed to download company logo from %s - keeping existing logo.", logo_url, exc_info=True)

    company.write(values)

    settings_values = {}
    for key, field in ACCOUNTING_FEATURE_FIELDS.items():
        parsed = _parse_bool(data.get(key))
        if parsed is not None:
            settings_values[field] = parsed
    if settings_values:
        _apply_config_settings(env, company, values, settings_values)


def _apply_config_settings(env, company, company_values, settings_values):
    """A module_* setting (e.g. budget_management -> module_account_budget)
    goes through Odoo's live "hot install" path, which takes an exclusive
    lock on ir_cron as a safety check - this can lose a race against a
    concurrently running server's own cron-polling thread and abort with a
    SerializationFailure. That's a textbook case for "just retry the
    transaction" - rolling back also discards the company.write() above
    (same uncommitted transaction), so redo it before each retry.
    """
    for attempt in range(3):
        try:
            env["res.config.settings"].create(settings_values).execute()
            return
        except SerializationFailure:
            if attempt == 2:
                raise
            env.cr.rollback()
            company.write(company_values)
            time.sleep(1)


def _import_res_partner(env, wb):
    columns = ("id", "name", "email", "phone", "is_company", "street", "city", "state", "country_id", "ref")
    for row in _sheet_rows(wb, "res.partner", columns):
        country = env["res.country"]
        if row["country_id"]:
            country = country.search([("name", "=", row["country_id"])], limit=1)
        state = env["res.country.state"]
        if row["state"]:
            state = state.search([("name", "=", row["state"]), ("country_id", "=", country.id)], limit=1)
        values = {
            "name": row["name"],
            "email": row["email"],
            "phone": row["phone"],
            "is_company": bool(row["is_company"]),
            "street": row["street"],
            "city": row["city"],
            "state_id": state.id,
            "country_id": country.id,
            "ref": row["ref"],
        }
        _get_or_create(env, "res.partner", row["id"], values)


def _import_kas_bank(env, wb, default_date):
    """Import kas/bank opening balances as account.bank.statement.line records
    so they show up as Bank Transactions (a plain account.move posted through
    the journal does not), matching how a manually-entered transaction works.
    Note: creating a statement line always posts it immediately (Odoo core
    behavior) - POST_OPENING_MOVES does not apply here.
    """
    company = env.ref("base.main_company")
    transfer_total = 0.0
    for move_data in _move_rows(wb, "kas_bank", default_date):
        journal = env["account.journal"].search([("name", "=", move_data["journal"])], limit=1)
        amount = sum(
            (line["debit"] or 0.0) - (line["credit"] or 0.0)
            for line in move_data["lines"]
            if _account_by_code_name(env, line["account"]) == journal.default_account_id
        )
        # The statement line's counterpart carries -amount on Liquidity
        # Transfer, so the opening balance needs +amount to net it out -
        # accumulate from the sheet regardless of whether this particular
        # line is (re)created below, so already-posted lines still count.
        transfer_total += amount

        existing = env.ref(move_data["id"], raise_if_not_found=False)
        if existing:
            if existing.state != "draft":
                continue
            # counterpart_account_id is a create()-only pseudo-field (Odoo
            # pops it before writing), so an existing draft line can't be
            # updated via write() - discard and recreate it instead, safe
            # since a draft statement line has no reconciliation yet.
            # unlink() also removes the underlying move; _get_or_create's
            # create path below reuses the now-stale ir.model.data row.
            existing.unlink()

        values = {
            "journal_id": journal.id,
            "date": move_data["date"],
            "payment_ref": move_data["ref"],
            "amount": amount,
            # Route the counterpart to the reconcilable "Liquidity Transfer"
            # account instead of the journal's (non-reconcilable) suspense
            # account, so it can be matched against the opening_debit/credit
            # counterpart entry on account.account.
            "counterpart_account_id": company.transfer_account_id.id,
        }
        _get_or_create(env, "account.bank.statement.line", move_data["id"], values)

    _write_opening_balances(env, {
        company.transfer_account_id: [max(transfer_total, 0.0), max(-transfer_total, 0.0)],
    })


def _import_account_account(env, wb):
    columns = ("id", "code", "name", "opening_debit", "opening_credit")
    opening_totals = {}
    for row in _sheet_rows(wb, "account.account", columns):
        account = env.ref(row["id"], raise_if_not_found=False)
        if not account:
            continue
        account.write({"code": row["code"], "name": row["name"]})
        if row["opening_debit"] or row["opening_credit"]:
            opening_totals[account] = [row["opening_debit"] or 0.0, row["opening_credit"] or 0.0]
    _write_opening_balances(env, opening_totals)


def _import_account_journal(env, wb):
    columns = ("id", "sequence", "name", "type", "code", "default_account_id", "Bank Feed")
    for row in _sheet_rows(wb, "account.journal", columns):
        values = {
            "sequence": row["sequence"],
            "name": row["name"],
            "type": JOURNAL_TYPES[row["type"]],
            "code": row["code"],
            "bank_statements_source": row["Bank Feed"],
        }
        account = _account_by_code_name(env, row["default_account_id"])
        if account:
            values["default_account_id"] = account.id
        _get_or_create(env, "account.journal", row["id"], values)


def _import_account_asset(env, wb, balance_date):
    columns = (
        "id", "name", "acquisition_date", "original_value", "already_depreciated_amount_import",
        "account_asset_id", "account_depreciation_id", "account_depreciation_expense_id",
        "method", "method_number", "method_period", "Jurnal",
    )
    opening_totals = {}  # account.account recordset -> [opening_debit, opening_credit]

    def _add_opening(account, debit=0.0, credit=0.0):
        if account:
            totals = opening_totals.setdefault(account, [0.0, 0.0])
            totals[0] += debit
            totals[1] += credit

    for row in _sheet_rows(wb, "account.asset", columns):
        journal = env["account.journal"].search([("name", "=", row["Jurnal"])], limit=1)
        asset_account = _account_by_code_name(env, row["account_asset_id"])
        depreciation_account = _account_by_code_name(env, row["account_depreciation_id"])
        method_period = ASSET_PERIODS[row["method_period"]]
        acquisition_date = row["acquisition_date"].date()

        already_depreciated = row["already_depreciated_amount_import"]
        if not already_depreciated:
            elapsed = _elapsed_depreciation_periods(acquisition_date, balance_date, method_period)
            elapsed = min(elapsed, row["method_number"])
            already_depreciated = (row["original_value"] or 0.0) / row["method_number"] * elapsed

        values = {
            "name": row["name"],
            "acquisition_date": acquisition_date,
            "original_value": row["original_value"],
            "already_depreciated_amount_import": already_depreciated,
            "account_asset_id": asset_account.id,
            "account_depreciation_id": depreciation_account.id,
            "account_depreciation_expense_id": _account_by_code_name(env, row["account_depreciation_expense_id"]).id,
            "method": ASSET_METHODS[row["method"]],
            "method_number": row["method_number"],
            "method_period": method_period,
            "journal_id": journal.id,
        }
        asset = _get_or_create(env, "account.asset", row["id"], values)
        if VALIDATE_IMPORTED_ASSETS and asset.state == "draft":
            asset.validate()

        _add_opening(asset_account, debit=row["original_value"] or 0.0)
        _add_opening(depreciation_account, credit=already_depreciated)

    _write_opening_balances(env, opening_totals)


def _reconcile_liquidity_transfer(env):
    """Auto-reconcile kas/bank opening entries against their Liquidity
    Transfer counterpart once both sides are posted (e.g. after the
    account.account opening_debit move has been reviewed and posted).
    """
    company = env.ref("base.main_company")
    lines = env["account.move.line"].search([
        ("account_id", "=", company.transfer_account_id.id),
        ("reconciled", "=", False),
        ("move_id.state", "=", "posted"),
    ])
    if lines:
        lines.reconcile()


def run_import(env, wb):
    """Run the full master-data import against an already-open workbook.
    Shared entry point for import_bundled_data() (a client module's
    post_init_hook) and the Settings > Import Data Master wizard (a
    user-uploaded file).
    """
    balance_date = _read_opening_balance_date(wb)
    # account_opening_date perlu tanggal *awal* tahun buku, bukan tanggal
    # cutover-nya - Odoo men-tanggal-kan opening move sehari sebelum awal
    # tahun buku, jadi arahnya kebalik dari balance_date.
    fiscal_year_start = balance_date + timedelta(days=1)
    if fiscal_year_start.day != 1:
        _logger.warning(
            'OPENING_BALANCE_DATE (%s) menghasilkan awal tahun buku %s, bukan tanggal 1 - '
            "cek lagi kalau ini bukan disengaja (mis. salah pilih tanggal).",
            balance_date, fiscal_year_start,
        )

    company = env.ref("base.main_company")
    company.write({"account_opening_date": fiscal_year_start})
    if company.account_opening_move_id and company.account_opening_move_id.state == "draft":
        company.account_opening_move_id.date = balance_date

    _import_company(env, wb)
    _import_res_partner(env, wb)
    _import_account_account(env, wb)
    _import_account_journal(env, wb)
    _import_account_asset(env, wb, balance_date)
    _import_kas_bank(env, wb, balance_date)
    _import_opening_move(env, wb, "vendor_bill", "supplier_rank", balance_date)
    _import_opening_move(env, wb, "customer_invoice", "customer_rank", balance_date)
    _reconcile_liquidity_transfer(env)
