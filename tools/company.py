"""Sheet company / c: profil perusahaan, pengaturan akuntansi, presisi analytic.
"""

import base64
import logging
import time

import requests
from psycopg2.errors import SerializationFailure

from .common import find_sheet, _parse_bool

_logger = logging.getLogger(__name__)


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


def _read_decimal_accuracy_percentage_analytic(wb):
    """Baca DECIMAL_ACCURACY_PERCENTAGE_ANALYTIC (opsional) dari sheet "company"
    (atau "c") - jumlah digit desimal yang dipakai widget alokasi analytic
    (field analytic_distribution, model decimal.precision "Percentage
    Analytic"). Beda project/klien bisa butuh presisi berbeda (atau tidak
    butuh sama sekali - default Odoo 2 digit sudah cukup untuk banyak kasus),
    makanya ini opsional dan dikontrol lewat sheet, bukan di-hardcode di
    modul - modul ini dipakai lintas project, bukan cuma satu klien.
    Return None kalau baris ini tidak ada di sheet manapun (tidak seperti
    _read_opening_balance_date, TIDAK raise - ini murni opsional).
    """
    for sheet_key in ("company", "c", "petunjuk", "p"):
        actual_sheet = find_sheet(wb, sheet_key)
        if not actual_sheet or actual_sheet not in wb.sheetnames:
            continue
        for row in wb[actual_sheet].iter_rows(values_only=True):
            if row and row[0] == "DECIMAL_ACCURACY_PERCENTAGE_ANALYTIC":
                return row[1]
    return None


def _apply_analytic_percentage_precision(env, digits, logger=None):
    """Naikkan presisi 'Percentage Analytic' (decimal.precision) sesuai
    DECIMAL_ACCURACY_PERCENTAGE_ANALYTIC di sheet, kalau diisi. Presisi 2
    digit (default Odoo) bisa bikin alokasi analytic dalam persen meleset
    ratusan rupiah dari nominal yang dimaksud (mis. 44.14% dari 9.062.900
    -> 4.000.364,06, padahal maunya tepat 4.000.000) karena rasio nominal
    yang diinginkan sering tidak "bulat" di 2 desimal.

    Sengaja TIDAK PERNAH menurunkan presisi yang sudah ada di database -
    record lain (project lain, atau import sebelumnya) mungkin sudah
    mengandalkan presisi yang lebih tinggi.
    """
    if digits in (None, ""):
        return

    try:
        digits = int(digits)
    except (TypeError, ValueError):
        msg = (
            f"DECIMAL_ACCURACY_PERCENTAGE_ANALYTIC harus berupa angka bulat, "
            f"ditemukan {digits!r} - diabaikan."
        )
        _logger.warning(msg)
        if logger:
            logger.log("company", 0, "DECIMAL_ACCURACY_PERCENTAGE_ANALYTIC", "", "warning", msg)
        return

    if not (0 <= digits <= 12):
        msg = (
            f"DECIMAL_ACCURACY_PERCENTAGE_ANALYTIC di luar rentang wajar (0-12): "
            f"{digits} - diabaikan."
        )
        _logger.warning(msg)
        if logger:
            logger.log("company", 0, "DECIMAL_ACCURACY_PERCENTAGE_ANALYTIC", "", "warning", msg)
        return

    precision = env["decimal.precision"].search([("name", "=", "Percentage Analytic")], limit=1)
    if not precision:
        return

    old_digits = precision.digits
    if digits <= old_digits:
        return

    precision.write({"digits": digits})
    msg = f"Presisi 'Percentage Analytic' dinaikkan dari {old_digits} ke {digits} digit."
    _logger.info(msg)
    if logger:
        logger.log("company", 0, "DECIMAL_ACCURACY_PERCENTAGE_ANALYTIC", "", "success", msg)


def _import_company(env, wb, logger=None):
    """Read the "company" / "c" sheet - key/value rows (also where
    OPENING_BALANCE_DATE lives, see _read_opening_balance_date) - and write
    it onto the single company record, then apply any accounting feature
    toggles found there (see
    ACCOUNTING_FEATURE_FIELDS). Optional sheet, and a blank cell leaves that
    field/toggle untouched rather than clearing it - so a partially-filled
    sheet (e.g. no logo URL yet) never blanks out data set another way.
    """
    sheet = find_sheet(wb, "company")
    if not sheet or sheet not in wb.sheetnames:
        return
    data = {}
    for row in wb[sheet].iter_rows(values_only=True):
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

    if logger:
        logger.log(
            sheet,
            1,
            str(company.id),
            company.name,
            "success",
            f"Profil dan konfigurasi perusahaan '{company.name}' berhasil diperbarui.",
        )
        logger.flush()


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
