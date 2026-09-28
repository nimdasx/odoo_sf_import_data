"""Sheet account.account / a.a (chart of accounts & saldo awal) dan
account.journal / a.j.
"""

from .common import (
    MODULE,
    find_sheet,
    _parse_bool,
    _sheet_rows,
    _get_or_create,
    _account_by_code_name,
    _write_opening_balances,
)


JOURNAL_TYPES = {"Bank": "bank", "Kas": "cash", "Lain-lain": "general"}


ACCOUNT_TYPES = {
    "piutang": "asset_receivable",
    "bank dan tunai": "asset_cash",
    "aktiva lancar": "asset_current",
    "aktiva tidak lancar": "asset_non_current",
    "prabayar": "asset_prepayments",
    "aktiva tetap": "asset_fixed",
    "utang": "liability_payable",
    "kartu kredit": "liability_credit_card",
    "pasiva terkini": "liability_current",
    "hutang tidak lancar": "liability_non_current",
    "ekuitas": "equity",
    "penghasilan tahun terkini": "equity_unaffected",
    "penghasilan": "income",
    "penghasilan lainnya": "income_other",
    "pengeluaran": "expense",
    "pengeluaran lainnya": "expense_other",
    "penyusutan": "expense_depreciation",
    "biaya pendapatan": "expense_direct_cost",
    "off-balance sheet": "off_balance",
}


def _import_account_account(env, wb, logger=None):
    sheet = find_sheet(wb, "account.account")
    if not sheet:
        return
    columns = ("id", "code", "name", "account_type", "opening_debit", "opening_credit", "active")
    opening_totals = {}
    company = env.ref("base.main_company")
    seen_codes = {}

    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
        code = str(row["code"]).strip() if row.get("code") is not None else ""
        name = str(row["name"]).strip() if row.get("name") is not None else ""
        if not code or not name:
            if logger:
                logger.log(sheet, row_num, code, name, "skipped", "Dilewati: Kode akun atau nama akun kosong.")
            continue

        is_duplicate = False
        if code in seen_codes:
            is_duplicate = True
            prev_row = seen_codes[code]
            if logger:
                logger.log(
                    sheet, row_num, code, name, "warning",
                    f"Peringatan: Kode akun '{code}' duplikat dari baris {prev_row}. Meng-update akun yang sudah ada.",
                )
        else:
            seen_codes[code] = row_num

        raw_type = row.get("account_type")
        account_type = None
        if raw_type:
            raw_str = str(raw_type).strip().lower()
            account_type = ACCOUNT_TYPES.get(raw_str) or (raw_str if raw_str in ACCOUNT_TYPES.values() else None)

        values = {"code": code, "name": name}
        if account_type:
            values["account_type"] = account_type

        if row.get("active") is not None:
            parsed_active = _parse_bool(row["active"])
            if parsed_active is not None:
                values["active"] = parsed_active

        account = False
        if code:
            account = env["account.account"].with_context(active_test=False).search([
                ("code", "=", code),
                ("company_ids", "in", company.id),
            ], limit=1)
            if not account:
                account = env["account.account"].with_context(active_test=False).search([("code", "=", code)], limit=1)
        if not account and row.get("id"):
            cand = env.ref(row["id"], raise_if_not_found=False)
            if cand and (not cand.code or cand.code == code):
                account = cand

        if account:
            account.write(values)
            if row.get("id"):
                if "." in row["id"]:
                    module, xml_name = row["id"].split(".", 1)
                else:
                    module, xml_name = MODULE, row["id"]
                data = env["ir.model.data"].search([("module", "=", module), ("name", "=", xml_name)])
                if not data:
                    env["ir.model.data"].create({"module": module, "name": xml_name, "model": "account.account", "res_id": account.id, "noupdate": True})
                else:
                    data.write({"noupdate": True, "res_id": account.id})
        else:
            if not values.get("account_type"):
                values["account_type"] = "asset_current"
            account_id = row.get("id") or f"account_account_{code}"
            account = _get_or_create(env, "account.account", account_id, values)

        if row.get("opening_debit") or row.get("opening_credit"):
            opening_totals[account] = [row["opening_debit"] or 0.0, row["opening_credit"] or 0.0]

        if logger and not is_duplicate:
            logger.log(sheet, row_num, code, name, "success", f"Akun [{code}] {name} berhasil diimport.")

    _write_opening_balances(env, opening_totals)
    if logger:
        logger.flush()


def _import_account_journal(env, wb, logger=None):
    sheet = find_sheet(wb, "account.journal")
    if not sheet:
        return
    columns = ("id", "sequence", "name", "type", "code", "default_account_id", "Bank Feed")
    company = env.ref("base.main_company")
    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
        code = str(row["code"]).strip() if row.get("code") is not None else ""
        name = str(row["name"]).strip() if row.get("name") is not None else ""
        if not code or not name:
            if logger:
                logger.log(sheet, row_num, code, name, "skipped", "Dilewati: Kode atau nama jurnal kosong.")
            continue

        journal_type = JOURNAL_TYPES.get(row.get("type"))
        if not journal_type:
            if logger:
                logger.log(
                    sheet, row_num, code, name, "skipped",
                    f"Dilewati: kolom type ({row.get('type')}) tidak dikenali. "
                    f"Nilai valid: {', '.join(JOURNAL_TYPES)}.",
                )
            continue

        values = {
            "sequence": row["sequence"],
            "name": row["name"],
            "type": journal_type,
            "code": row["code"],
            "bank_statements_source": row["Bank Feed"],
        }
        account = _account_by_code_name(env, row["default_account_id"])
        if account:
            values["default_account_id"] = account.id

        # Search by code within company first to respect UNIQUE(company_id, code) constraint
        journal = False
        if row.get("code"):
            journal = env["account.journal"].with_context(active_test=False).search([
                ("code", "=", str(row["code"]).strip()),
                ("company_id", "=", company.id),
            ], limit=1)
        if not journal and row.get("id"):
            journal = env.ref(row["id"], raise_if_not_found=False)
        if not journal and row.get("name"):
            journal = env["account.journal"].with_context(active_test=False).search([
                ("name", "=", str(row["name"]).strip()),
                ("company_id", "=", company.id),
            ], limit=1)

        if journal:
            journal.write(values)
            if row.get("id"):
                if "." in row["id"]:
                    module, xml_name = row["id"].split(".", 1)
                else:
                    module, xml_name = MODULE, row["id"]
                data = env["ir.model.data"].search([("module", "=", module), ("name", "=", xml_name)])
                if not data:
                    env["ir.model.data"].create({"module": module, "name": xml_name, "model": "account.journal", "res_id": journal.id, "noupdate": True})
                else:
                    data.write({"noupdate": True, "res_id": journal.id})
        else:
            journal_id = row.get("id") or f"account_journal_{code.lower()}"
            journal = _get_or_create(env, "account.journal", journal_id, values)

        if logger:
            acc_info = f" (Akun: {account.code} {account.name})" if account else ""
            logger.log(sheet, row_num, code, name, "success", f"Jurnal [{code}] {name} berhasil disinkronkan{acc_info}.")

    if logger:
        logger.flush()
