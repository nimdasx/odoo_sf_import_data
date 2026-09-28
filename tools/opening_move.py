"""Opening move partner dari sheet vendor_bill / v.b & customer_invoice / c.i,
saldo awal kas/bank, dan reconcile liquidity transfer.
"""

import logging

from odoo.fields import Command

from .common import (
    MODULE,
    find_sheet,
    POST_OPENING_MOVES,
    _parse_sheet_date,
    _sheet_rows,
    _get_or_create,
    _account_by_code_name,
    _write_opening_balances,
)

_logger = logging.getLogger(__name__)


def _partner_by_name(env, name, rank_field):
    if not name:
        return env["res.partner"]
    partner = env["res.partner"].search([("name", "=", name)], limit=1)
    return partner or env["res.partner"].create({"name": name, rank_field: 1})


def _move_rows(wb, sheet, default_date):
    """Group an Odoo-import-style sheet (a header row per move, followed by
    blank-id continuation rows) into one dict per move with a "lines" list.
    """
    actual_sheet = find_sheet(wb, sheet)
    if not actual_sheet or actual_sheet not in wb.sheetnames:
        return
    rows = wb[actual_sheet].iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        return
    if not header:
        return
    idx = {col: header.index(col) for col in header}

    def get(row, col):
        i = idx.get(col)
        return row[i] if i is not None and i < len(row) else None

    move = None
    for row_num, row in enumerate(rows, start=2):
        xml_id = get(row, "id")
        if xml_id:
            if move:
                yield move
            xml_id = xml_id if "." in xml_id else f"{MODULE}.{xml_id}"
            raw_date = get(row, "date")
            move = {
                "id": xml_id,
                "_row_number": row_num,
                "ref": get(row, "Reference"),
                "date": _parse_sheet_date(raw_date, "date") if raw_date else default_date,
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


def _import_opening_move(env, wb, sheet, partner_rank_field, default_date, logger=None):
    actual_sheet = find_sheet(wb, sheet)
    if not actual_sheet:
        return
    company = env.ref("base.main_company")
    for move_data in _move_rows(wb, actual_sheet, default_date):
        row_num = move_data.get("_row_number", 0)
        existing = env.ref(move_data["id"], raise_if_not_found=False)
        if existing and existing.state != "draft":
            if logger:
                logger.log(
                    actual_sheet,
                    row_num,
                    move_data.get("id", ""),
                    move_data.get("ref", ""),
                    "skipped",
                    f"Transaksi {move_data.get('ref')} sudah berstatus posted - dilewati.",
                )
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
                vals["date_maturity"] = _parse_sheet_date(line["date_maturity"], "date_maturity")
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

        if logger:
            logger.log(
                actual_sheet,
                row_num,
                move_data.get("id", ""),
                move_data.get("ref", "") or move_data.get("id", ""),
                "success",
                f"Transaksi {actual_sheet} [{move_data.get('ref')}] berhasil diimport ({len(move_data['lines'])} baris detail).",
            )
    if logger:
        logger.flush()


def _import_kas_bank(env, wb, default_date, logger=None):
    """Import kas/bank opening balances as account.bank.statement.line records.
    Source: opening_balance column on sheet "account.journal" / "a.j".
    """
    company = env.ref("base.main_company")
    transfer_total = 0.0

    journal_sheet = find_sheet(wb, "account.journal")
    if journal_sheet:
        columns = ("id", "name", "type", "code", "opening_balance")
        rows = list(_sheet_rows(wb, journal_sheet, columns))
        if any(r.get("opening_balance") not in (None, "") for r in rows):

            # Clean up all existing opening statement lines on default_date for bank/cash journals
            all_bank_cash = env["account.journal"].search([
                ("type", "in", ("bank", "cash")),
                ("company_id", "=", company.id),
            ])
            existing_lines = env["account.bank.statement.line"].search([
                ("journal_id", "in", all_bank_cash.ids),
                ("date", "=", default_date),
            ])
            for old in existing_lines:
                if old.is_reconciled:
                    old.line_ids.remove_move_reconcile()
                old.unlink()

            for row in rows:
                raw_bal = row.get("opening_balance")
                if raw_bal in (None, ""):
                    continue
                try:
                    amount = float(raw_bal)
                except (ValueError, TypeError):
                    _logger.warning("Nilai opening_balance %r pada jurnal %r tidak valid - dilewati.", raw_bal, row.get("name"))
                    continue

                journal = env.ref(row["id"], raise_if_not_found=False)
                if not journal and row.get("code"):
                    journal = env["account.journal"].search([
                        ("code", "=", str(row["code"]).strip()),
                        ("company_id", "=", company.id),
                    ], limit=1)
                if not journal and row.get("name"):
                    journal = env["account.journal"].search([
                        ("name", "=", str(row["name"]).strip()),
                        ("company_id", "=", company.id),
                    ], limit=1)

                if not journal:
                    _logger.warning("Jurnal %r tidak ditemukan untuk saldo awal kas/bank - dilewati.", row.get("name"))
                    continue

                transfer_total += amount
                line_xml_id = f"{MODULE}.statement_line_{row['id'].replace('.', '_')}"

                values = {
                    "journal_id": journal.id,
                    "date": default_date,
                    "payment_ref": "Saldo Awal",
                    "amount": amount,
                    "counterpart_account_id": company.transfer_account_id.id,
                }
                _get_or_create(env, "account.bank.statement.line", line_xml_id, values)
                if logger:
                    logger.log(
                        f"{journal_sheet} (Saldo Awal)",
                        row.get("_row_number", 0),
                        journal.code,
                        journal.name,
                        "success",
                        f"Saldo awal Kas/Bank [{journal.code}] {journal.name} sebesar Rp {amount:,.2f} berhasil dicatat.",
                    )
            if logger:
                logger.flush()

    _write_opening_balances(env, {
        company.transfer_account_id: [max(transfer_total, 0.0), max(-transfer_total, 0.0)],
    })


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
