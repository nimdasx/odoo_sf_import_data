"""Sheet res.partner / r.p: kontak, akun default hutang/piutang, dan saldo
awal hutang/piutang per partner.
"""

from odoo.fields import Command

from .common import (
    MODULE,
    find_sheet,
    POST_OPENING_MOVES,
    _parse_sheet_date,
    _sheet_rows,
    _get_or_create,
    _account_by_code_name,
)
from .opening_move import _move_rows


def _import_res_partner(env, wb, logger=None):
    sheet = find_sheet(wb, "res.partner")
    if not sheet:
        return
    columns = (
        "id", "name", "email", "phone", "is_company", "street", "city", "state", "country_id", "ref",
        "customer_rank", "supplier_rank",
    )
    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
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
        # Only written when filled - a blank cell must not reset a rank Odoo
        # already bumped from posted invoices/bills on re-import.
        for rank_field in ("customer_rank", "supplier_rank"):
            if row[rank_field] is not None:
                values[rank_field] = int(row[rank_field])
        partner_id = row.get("id") or f"partner_{row['name']}"
        _get_or_create(env, "res.partner", partner_id, values)
        if logger:
            logger.log(
                sheet,
                row_num,
                str(row.get("ref") or row.get("id", "") or ""),
                str(row.get("name", "")),
                "success",
                f"Kontak / Partner '{row['name']}' berhasil disinkronkan.",
            )
    if logger:
        logger.flush()


# (kolom / field akun default partner, tipe akun yang diterima Odoo, label).
PARTNER_DEFAULT_ACCOUNTS = (
    ("property_account_payable_id", "liability_payable", "hutang"),
    ("property_account_receivable_id", "asset_receivable", "piutang"),
)


def _import_partner_default_accounts(env, wb, logger=None):
    """Isi akun hutang/piutang default partner dari kolom property_account_*
    di sheet r.p - kolom yang sama dipakai _import_partner_opening_balance.
    Terpisah dari _import_res_partner karena harus jalan setelah
    _import_account_account (di database baru akunnya belum ada). Sel kosong
    tidak mengubah akun default yang sudah ada.
    """
    sheet = find_sheet(wb, "res.partner")
    if not sheet:
        return
    columns = ("id", "name") + tuple(field for field, _type, _label in PARTNER_DEFAULT_ACCOUNTS)
    company = env.ref("base.main_company")
    for row in _sheet_rows(wb, sheet, columns):
        if not any(row.get(field) for field, _type, _label in PARTNER_DEFAULT_ACCOUNTS):
            continue
        row_num = row.get("_row_number", 0)
        identifier = row.get("id") or f"Row {row_num}"
        partner = env.ref(row["id"], raise_if_not_found=False) if row.get("id") else None
        if not partner and row.get("name"):
            partner = env["res.partner"].search([("name", "=", row["name"])], limit=1)
        if not partner:
            if logger:
                logger.log(
                    sheet, row_num, identifier, row.get("name", ""), "warning",
                    f"Akun default '{row.get('name')}' dilewati: partner tidak ditemukan.",
                )
            continue

        values, warnings = {}, []
        for field, account_type, label in PARTNER_DEFAULT_ACCOUNTS:
            if not row.get(field):
                continue
            account = _account_by_code_name(env, row[field])
            # Akun default partner dipakai invoice/bill baru, dan Odoo menolak
            # baris invoice/bill di akun hutang/piutang yang tipenya tidak
            # cocok - jadi di sini lebih ketat dari saldo awal (yang cukup warning).
            if not account:
                warnings.append(f"akun {label} ({row[field]}) tidak ditemukan")
            elif not account.active:
                warnings.append(f"akun {label} {account.code} nonaktif")
            elif account.account_type != account_type:
                warnings.append(f"akun {label} {account.code} bukan tipe {account_type}")
            else:
                values[field] = account.id
        if values:
            partner.with_company(company).write(values)
        if logger:
            message = f"Akun default hutang/piutang '{partner.name}'"
            if values:
                message += " diisi: " + ", ".join(
                    env["account.account"].browse(account_id).code for account_id in values.values()
                )
            if warnings:
                message += (" - " if values else " tidak diisi: ") + "; ".join(warnings)
            logger.log(
                sheet, row_num, identifier, row.get("name", ""),
                "warning" if warnings else "success", message + ".",
            )
    if logger:
        logger.flush()


# Saldo awal hutang/piutang ringkas per partner, diisi langsung di sheet
# res.partner / r.p (alternatif yang lebih mudah dari v.b / c.i). Per jenis:
# (label, kolom akun, kolom nominal, kolom jatuh tempo, kolom referensi,
#  field akun default partner, rank partner, xml_id journal entry, sisi normal).
# Kolom akun sengaja dinamai sama dengan field akun default partner.
PARTNER_OPENING_KINDS = (
    ("Hutang", "property_account_payable_id", "nominal_hutang", "jatuh_tempo_hutang", "ref_hutang",
     "property_account_payable_id", "supplier_rank", "opening_payable_partner", "credit"),
    ("Piutang", "property_account_receivable_id", "nominal_piutang", "jatuh_tempo_piutang", "ref_piutang",
     "property_account_receivable_id", "customer_rank", "opening_receivable_partner", "debit"),
)


def _opening_journal(env, company):
    """Jurnal yang sama dengan opening move perusahaan; kalau opening move
    belum ada, cari dengan cara yang sama seperti Odoo membuatnya
    (res.company._get_default_opening_move_values): jurnal umum pertama.
    """
    if company.account_opening_move_id:
        return company.account_opening_move_id.journal_id
    return env["account.journal"].search([
        *env["account.journal"]._check_company_domain(company),
        ("type", "=", "general"),
    ], limit=1)


def _opening_move_partner_accounts(wb, default_date):
    """(nama partner, kode akun) yang sudah diisi di sheet v.b / c.i - untuk
    memperingatkan saldo yang juga diisi di r.p (tercatat ganda).
    """
    pairs = set()
    for sheet in ("vendor_bill", "customer_invoice"):
        actual_sheet = find_sheet(wb, sheet)
        if not actual_sheet:
            continue
        for move_data in _move_rows(wb, actual_sheet, default_date):
            for line in move_data["lines"]:
                if line["partner"]:
                    pairs.add((str(line["partner"]).strip(), str(line["account"]).split(" ", 1)[0]))
    return pairs


def _import_partner_opening_balance(env, wb, balance_date, logger=None):
    """Satu journal entry untuk seluruh saldo awal hutang dan satu untuk
    seluruh saldo awal piutang dari kolom tambahan di sheet r.p - satu baris
    jurnal per partner (tetap bisa di-reconcile & masuk aged report per
    partner), plus satu baris penyeimbang ke akun unaffected earnings seperti
    v.b / c.i. Baris yang bermasalah dilewati per partner, bukan
    menggagalkan seluruh journal entry.
    """
    sheet = find_sheet(wb, "res.partner")
    if not sheet:
        return
    columns = ("id", "name", "customer_rank", "supplier_rank")
    for kind in PARTNER_OPENING_KINDS:
        columns += kind[1:5]
    rows = list(_sheet_rows(wb, sheet, columns))
    company = env.ref("base.main_company")
    journal = _opening_journal(env, company)
    move_partner_accounts = _opening_move_partner_accounts(wb, balance_date)

    for (label, account_col, amount_col, maturity_col, ref_col,
         default_account_field, rank_field, move_xml_name, normal_side) in PARTNER_OPENING_KINDS:
        log_sheet = f"{sheet} ({label})"
        move_xml_id = f"{MODULE}.{move_xml_name}"
        existing = env.ref(move_xml_id, raise_if_not_found=False)
        if existing and existing.state != "draft":
            if any(row.get(amount_col) not in (None, "", 0) for row in rows) and logger:
                logger.log(
                    log_sheet, 0, move_xml_name, existing.ref, "skipped",
                    f"Journal entry saldo awal {label.lower()} partner sudah berstatus posted - dilewati. "
                    "Koreksi lewat jurnal penyesuaian terpisah.",
                )
            continue

        line_commands = []
        balance = 0.0
        for row in rows:
            amount = row.get(amount_col)
            if amount in (None, "", 0):
                continue
            row_num = row.get("_row_number", 0)
            identifier = row.get("id") or f"Row {row_num}"

            def _skip(reason):
                if logger:
                    logger.log(
                        log_sheet, row_num, identifier, row.get("name", ""), "warning",
                        f"Saldo awal {label.lower()} '{row.get('name')}' dilewati: {reason}.",
                    )

            if not isinstance(amount, (int, float)):
                _skip(f"{amount_col} '{amount}' bukan angka")
                continue
            partner = env.ref(row["id"], raise_if_not_found=False) if row.get("id") else None
            if not partner and row.get("name"):
                partner = env["res.partner"].search([("name", "=", row["name"])], limit=1)
            if not partner:
                _skip("partner tidak ditemukan")
                continue

            if row.get(account_col):
                account = _account_by_code_name(env, row[account_col])
                if not account:
                    _skip(f"akun ({row[account_col]}) tidak ditemukan")
                    continue
            else:
                account = partner.with_company(company)[default_account_field]
                if not account:
                    _skip(f"{account_col} kosong dan partner tidak punya akun {label.lower()} default")
                    continue
            # Akun default bawaan chart of account sering sudah dinonaktifkan
            # oleh _cleanup_previous_data (tidak ada di sheet a.a).
            if not account.active:
                _skip(f"akun {account.code} nonaktif - isi {account_col} dengan akun aktif dari sheet a.a")
                continue

            try:
                date_maturity = (
                    _parse_sheet_date(row[maturity_col], maturity_col)
                    if row.get(maturity_col) not in (None, "")
                    else balance_date
                )
            except ValueError as e:
                _skip(str(e))
                continue

            # Nominal positif = saldo normal (hutang di kredit, piutang di
            # debit); negatif = saldo terbalik, mis. uang muka ke vendor.
            debit, credit = (amount, 0.0) if normal_side == "debit" else (0.0, amount)
            if amount < 0:
                debit, credit = -credit, -debit
            balance += credit - debit
            name = f"Saldo awal {label.lower()} - {partner.name}"
            if row.get(ref_col) not in (None, ""):
                name += f" ({row[ref_col]})"
            line_commands.append(Command.create({
                "account_id": account.id,
                "partner_id": partner.id,
                "debit": debit,
                "credit": credit,
                "name": name,
                "date_maturity": date_maturity,
            }))

            # Sama seperti kolom rank di r.p: hanya diisi kalau kolomnya
            # kosong, supaya nilai eksplisit di sheet tidak ditimpa.
            if row.get(rank_field) in (None, "") and not partner[rank_field]:
                partner[rank_field] = 1

            if logger:
                warnings = []
                if account.account_type not in ("liability_payable", "asset_receivable"):
                    warnings.append(
                        f"akun {account.code} bukan tipe hutang/piutang - tidak muncul di Aged Payable/Receivable"
                    )
                if (partner.name.strip(), account.code) in move_partner_accounts:
                    warnings.append(
                        f"partner dengan akun {account.code} juga diisi di sheet v.b / c.i - cek supaya tidak tercatat ganda"
                    )
                logger.log(
                    log_sheet, row_num, identifier, row.get("name", ""),
                    "warning" if warnings else "success",
                    f"Saldo awal {label.lower()} '{partner.name}' Rp {amount:,.2f} berhasil diimport"
                    + (f" ({'; '.join(warnings)})" if warnings else "") + ".",
                )

        if not line_commands:
            # Semua nominal dikosongkan di import ulang - buang draft lama.
            if existing:
                existing.unlink()
            continue

        if not company.currency_id.is_zero(balance):
            line_commands.append(Command.create({
                "account_id": company.get_unaffected_earnings_account().id,
                "debit": max(balance, 0.0),
                "credit": max(-balance, 0.0),
                "name": "Automatic Balancing Line",
            }))
        values = {
            "ref": f"Saldo awal {label.lower()} partner",
            "date": balance_date,
            "journal_id": journal.id,
            "move_type": "entry",
            "line_ids": [Command.clear()] + line_commands,
        }
        move = _get_or_create(env, "account.move", move_xml_id, values)
        if POST_OPENING_MOVES and move.state == "draft":
            move.action_post()
    if logger:
        logger.flush()
