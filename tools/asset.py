"""Sheet account.asset.model / a.a.m dan account.asset / a.as, termasuk
hitungan akumulasi penyusutan untuk opening balance.
"""

import calendar
import logging

from dateutil.relativedelta import relativedelta

from .common import (
    MODULE,
    find_sheet,
    _parse_sheet_date,
    _sheet_rows,
    _get_or_create,
    _account_by_code_name,
    _write_opening_balances,
)

_logger = logging.getLogger(__name__)


ASSET_METHODS = {"Garis Lurus": "linear"}


ASSET_PERIODS = {"Bulan": "1", "Tahun": "12"}


ASSET_PRORATA_COMPUTATION_TYPES = ("none", "constant_periods", "daily_computation")


# Set to True to auto-validate imported assets (generates the depreciation
# board and posts its moves). False leaves them as draft for manual review.
VALIDATE_IMPORTED_ASSETS = False


def _elapsed_depreciation_periods(acquisition_date, balance_date, method_period):
    """Prorata periods from acquisition_date through balance_date, matching
    Odoo 19 constant_periods prorata calculation convention.
    """
    if not acquisition_date or not balance_date or acquisition_date > balance_date:
        return 0.0

    days_in_acq_month = calendar.monthrange(acquisition_date.year, acquisition_date.month)[1]
    days_in_bal_month = calendar.monthrange(balance_date.year, balance_date.month)[1]

    start_prorata = (days_in_acq_month - acquisition_date.day + 1) / days_in_acq_month
    end_prorata = balance_date.day / days_in_bal_month

    elapsed_months = (
        start_prorata
        + end_prorata
        + (balance_date.year - acquisition_date.year) * 12
        + (balance_date.month - acquisition_date.month - 1)
    )

    if method_period == "12":  # yearly
        return max(0.0, elapsed_months / 12.0)
    return max(0.0, elapsed_months)


def _accumulated_depreciation(depreciable_value, acquisition_date, asof_date, method_number, method_period):
    """Akumulasi penyusutan (metode constant_periods Odoo 19) dari acquisition_date
    sampai asof_date - dipakai untuk already_depreciated_amount_import (asof_date
    = balance_date). depreciable_value = original_value - salvage_value, sama
    seperti total_depreciable_value di account.asset.
    """
    if not method_number:
        return 0.0
    elapsed = _elapsed_depreciation_periods(acquisition_date, asof_date, method_period)
    elapsed = min(elapsed, method_number)
    return depreciable_value / method_number * elapsed


def _mid_month_prorata_date(acquisition_date):
    """Konvensi pertengahan bulan: perolehan sebelum tanggal 15 mulai disusutkan
    dari tanggal 1 bulan itu, tanggal 15 ke atas dari tanggal 1 bulan berikutnya.
    """
    first_of_month = acquisition_date.replace(day=1)
    if acquisition_date.day < 15:
        return first_of_month
    return first_of_month + relativedelta(months=1)


def _import_account_asset_model(env, wb, logger=None):
    sheet = find_sheet(wb, "account.asset.model")
    if not sheet or "account.asset" not in env:
        return
    columns = (
        "id", "name", "account_asset_id", "account_depreciation_id",
        "account_depreciation_expense_id", "method", "method_number", "method_period", "Jurnal",
    )
    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
        if not row.get("name"):
            continue

        raw_period = row.get("method_period")
        method_period = ASSET_PERIODS.get(raw_period, "1")
        method = ASSET_METHODS.get(row.get("method"), "linear")
        method_number = int(row.get("method_number") or 1)

        asset_account = _account_by_code_name(env, row.get("account_asset_id"))
        depreciation_account = _account_by_code_name(env, row.get("account_depreciation_id"))
        dep_expense_account = _account_by_code_name(env, row.get("account_depreciation_expense_id"))

        journal = env["account.journal"].search([("name", "=", row.get("Jurnal"))], limit=1)
        if not journal:
            journal = env["account.journal"].search([("type", "=", "general")], limit=1)

        values = {
            "name": row["name"],
            "state": "model",
            "account_asset_id": asset_account.id if asset_account else False,
            "account_depreciation_id": depreciation_account.id if depreciation_account else False,
            "account_depreciation_expense_id": dep_expense_account.id if dep_expense_account else False,
            "method": method,
            "method_number": method_number,
            "method_period": method_period,
            "journal_id": journal.id if journal else False,
        }

        model_id = row.get("id") or f"account_asset_model_{row['name'].lower().replace(' ', '_')}"
        model = False
        if row.get("id"):
            model = env.ref(row["id"], raise_if_not_found=False)
        if not model:
            model = env["account.asset"].search([("name", "=", row["name"]), ("state", "=", "model")], limit=1)

        if model:
            model.write(values)
            if row.get("id"):
                if "." in row["id"]:
                    module, xml_name = row["id"].split(".", 1)
                else:
                    module, xml_name = MODULE, row["id"]
                data = env["ir.model.data"].search([("module", "=", module), ("name", "=", xml_name)])
                if not data:
                    env["ir.model.data"].create({"module": module, "name": xml_name, "model": "account.asset", "res_id": model.id, "noupdate": True})
                else:
                    data.write({"noupdate": True, "res_id": model.id})
        else:
            model = _get_or_create(env, "account.asset", model_id, values)

        if logger:
            logger.log(
                sheet,
                row_num,
                row.get("id", ""),
                row.get("name", ""),
                "success",
                f"Model Aset '{row['name']}' berhasil disinkronkan.",
            )

    if logger:
        logger.flush()


def _import_account_asset(env, wb, balance_date, logger=None):
    sheet = find_sheet(wb, "account.asset")
    if not sheet:
        return
    columns = (
        "id", "name", "acquisition_date", "original_value", "already_depreciated_amount_import",
        "account_asset_id", "account_depreciation_id", "account_depreciation_expense_id",
        "method", "method_number", "method_period", "Jurnal", "prorata_computation_type",
        "salvage_value",
    )
    opening_totals = {}  # account.account recordset -> [opening_debit, opening_credit]

    def _add_opening(account, debit=0.0, credit=0.0):
        if account:
            totals = opening_totals.setdefault(account, [0.0, 0.0])
            totals[0] += debit
            totals[1] += credit

    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
        if not row.get("name"):
            continue

        raw_period = row.get("method_period")
        if not raw_period or raw_period not in ASSET_PERIODS:
            if logger:
                logger.log(
                    sheet,
                    row_num,
                    row.get("id", "") or f"Row {row_num}",
                    row.get("name", ""),
                    "skipped",
                    f"Aset '{row['name']}' dilewati: kolom periode metode ({raw_period}) tidak terisi / tidak valid.",
                )
            continue

        asset_account = _account_by_code_name(env, row.get("account_asset_id"))
        if not asset_account:
            if logger:
                logger.log(
                    sheet,
                    row_num,
                    row.get("id", "") or f"Row {row_num}",
                    row.get("name", ""),
                    "skipped",
                    f"Aset '{row['name']}' dilewati: akun aset ({row.get('account_asset_id')}) tidak ditemukan.",
                )
            continue

        journal = env["account.journal"].search([("name", "=", row.get("Jurnal"))], limit=1)
        if not journal:
            journal = env["account.journal"].search([("type", "=", "general")], limit=1)

        depreciation_account = _account_by_code_name(env, row.get("account_depreciation_id"))
        dep_expense_account = _account_by_code_name(env, row.get("account_depreciation_expense_id"))
        method_period = ASSET_PERIODS[raw_period]
        # Aset tanpa tanggal perolehan tidak bisa dihitung penyusutannya, dan
        # aset yang diperoleh setelah OPENING_BALANCE_DATE belum dimiliki per
        # tanggal saldo awal (pembeliannya dicatat lewat transaksi biasa) -
        # keduanya dilewati supaya tidak ikut masuk opening balance. Sengaja
        # dibandingkan dengan acquisition_date, bukan prorata_date: aset yang
        # dibeli 15-31 Desember sudah dimiliki per 31 Desember walau
        # penyusutannya baru mulai 1 Januari.
        raw_acquisition_date = row.get("acquisition_date")
        acquisition_date = (
            _parse_sheet_date(raw_acquisition_date, "acquisition_date")
            if raw_acquisition_date not in (None, "")
            else None
        )
        if not acquisition_date or acquisition_date > balance_date:
            if logger:
                reason = (
                    "tanggal perolehan (acquisition_date) kosong"
                    if not acquisition_date
                    else f"tanggal perolehan {acquisition_date} setelah OPENING_BALANCE_DATE {balance_date}"
                )
                logger.log(
                    sheet,
                    row_num,
                    row.get("id", "") or f"Row {row_num}",
                    row.get("name", ""),
                    "warning",
                    f"Aset '{row['name']}' dilewati: {reason}.",
                )
            continue

        original_value = row.get("original_value") or 0.0
        method_number = int(row.get("method_number") or 1)
        method = ASSET_METHODS.get(row.get("method"), "linear")

        # Nilai residu (Not Depreciable Value di Odoo) tidak ikut disusutkan,
        # jadi akumulasi otomatis di bawah dihitung dari original_value -
        # salvage_value supaya cocok dengan depreciation board Odoo.
        salvage_value = row.get("salvage_value") or 0.0
        if not isinstance(salvage_value, (int, float)) or not 0 <= salvage_value <= original_value:
            if logger:
                logger.log(
                    sheet,
                    row_num,
                    row.get("id", "") or f"Row {row_num}",
                    row.get("name", ""),
                    "warning",
                    f"Aset '{row['name']}': salvage_value '{row.get('salvage_value')}' tidak valid "
                    f"(harus angka 0 s.d. original_value {original_value:,.2f}) - dianggap 0.",
                )
            salvage_value = 0.0

        prorata_computation_type = str(row.get("prorata_computation_type") or "").strip().lower() or False
        if prorata_computation_type and prorata_computation_type not in ASSET_PRORATA_COMPUTATION_TYPES:
            if logger:
                logger.log(
                    sheet,
                    row_num,
                    row.get("id", "") or f"Row {row_num}",
                    row.get("name", ""),
                    "warning",
                    f"Aset '{row['name']}': prorata_computation_type '{row.get('prorata_computation_type')}' "
                    f"tidak valid (pilihan: {', '.join(ASSET_PRORATA_COMPUTATION_TYPES)}) - diabaikan.",
                )
            prorata_computation_type = False
        prorata_date = None
        if prorata_computation_type == "constant_periods":
            prorata_date = _mid_month_prorata_date(acquisition_date)
        # Odoo memulai jadwal penyusutan dari prorata_date, jadi akumulasi
        # otomatis di bawah juga harus dihitung dari tanggal yang sama supaya
        # opening balance cocok dengan depreciation board-nya.
        depreciation_start = prorata_date or acquisition_date

        already_depreciated = row.get("already_depreciated_amount_import")
        if already_depreciated in (None, ""):
            already_depreciated = _accumulated_depreciation(
                original_value - salvage_value, depreciation_start, balance_date, method_number, method_period
            )

        values = {
            "name": row["name"],
            "acquisition_date": acquisition_date,
            "original_value": original_value,
            "salvage_value": salvage_value,
            "already_depreciated_amount_import": already_depreciated,
            "account_asset_id": asset_account.id,
            "account_depreciation_id": depreciation_account.id if depreciation_account else False,
            "account_depreciation_expense_id": dep_expense_account.id if dep_expense_account else False,
            "method": method,
            "method_number": method_number,
            "method_period": method_period,
            "journal_id": journal.id if journal else False,
        }
        if prorata_computation_type:
            values["prorata_computation_type"] = prorata_computation_type
        if prorata_date:
            values["prorata_date"] = prorata_date
        asset_id = row.get("id") or f"account_asset_{row['name']}"
        asset = _get_or_create(env, "account.asset", asset_id, values)
        if VALIDATE_IMPORTED_ASSETS and asset.state == "draft":
            try:
                asset.validate()
            except Exception as e:
                _logger.warning("Gagal validasi aset %s: %s", asset.name, e)

        _add_opening(asset_account, debit=original_value)
        if depreciation_account:
            _add_opening(depreciation_account, credit=already_depreciated)

        if logger:
            logger.log(
                sheet,
                row_num,
                row.get("id", ""),
                row.get("name", ""),
                "success",
                f"Aset '{row['name']}' nilai perolehan Rp {original_value:,.2f} berhasil diimport.",
            )

    _write_opening_balances(env, opening_totals)
    if logger:
        logger.flush()
