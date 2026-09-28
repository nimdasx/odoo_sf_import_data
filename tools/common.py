"""Helper bersama importer: alias sheet, pembacaan baris sheet, parser sel,
logger import, dan penulisan opening balance.
"""

import logging
import os
from datetime import date, datetime

_logger = logging.getLogger(__name__)


# New xml_ids created by this import (e.g. for a journal/asset row that has
# no existing match) are registered under this addon's own name, regardless
# of which client module's data actually triggered the import.
MODULE = os.path.basename(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


SHEET_ALIASES = {
    "petunjuk": ("petunjuk", "p"),
    "company": ("company", "c"),
    "res.partner": ("res.partner", "r.p"),
    "account.account": ("account.account", "a.a"),
    "account.journal": ("account.journal", "a.j"),
    "account.asset": ("account.asset", "a.as"),
    "account.asset.model": ("account.asset.model", "a.a.m"),
    "vendor_bill": ("vendor_bill", "v.b"),
    "customer_invoice": ("customer_invoice", "c.i"),
    "account.analytic.plan": ("account.analytic.plan", "a.an.p"),
    "account.analytic.account": ("account.analytic.account", "a.an.a"),
    "account.report": ("account.report", "a.r"),
    "account.report.line": ("account.report.line", "a.r.l"),
}


def find_sheet(wb, canonical_or_alias):
    """Cari nama sheet sesungguhnya yang ada pada workbook (wb.sheetnames).
    Mendukung nama kanonikal, nama alias singkat (misal 'a.a' untuk 'account.account'),
    serta toleransi case-insensitif dan penghapusan spasi.
    """
    if not wb or not hasattr(wb, "sheetnames"):
        return None
    sheetnames = wb.sheetnames
    if canonical_or_alias in sheetnames:
        return canonical_or_alias

    candidates = [canonical_or_alias]
    if canonical_or_alias in SHEET_ALIASES:
        candidates.extend(SHEET_ALIASES[canonical_or_alias])
    else:
        for canon, aliases in SHEET_ALIASES.items():
            if canonical_or_alias in aliases or canonical_or_alias == canon:
                candidates.append(canon)
                candidates.extend(aliases)
                break

    for cand in candidates:
        if cand in sheetnames:
            return cand

    clean_map = {s.strip().lower(): s for s in sheetnames}
    for cand in candidates:
        cand_clean = cand.strip().lower()
        if cand_clean in clean_map:
            return clean_map[cand_clean]

    return None


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


# Set to True to auto-post imported opening vendor/customer moves.
# False leaves them as draft for manual review.
POST_OPENING_MOVES = False


_DATE_STRING_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")


def _parse_sheet_date(value, context):
    """A date cell normally comes back from openpyxl as a real datetime,
    but if someone typed the date as plain text in Google Sheets instead
    of a proper Date-formatted cell, it comes back as a str instead -
    try a few common formats before giving up, so that crashes with a
    clear message instead of an opaque "'str' object has no attribute
    'date'" AttributeError.
    """
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in _DATE_STRING_FORMATS:
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    raise ValueError(
        f"{context}: nilai tanggal {value!r} tidak dikenali - pastikan cell-nya "
        "berformat Date (bukan Text) di sheet-nya."
    )


class ImportLogger:
    def __init__(self, history):
        self.history = history
        self.pending_lines = []

    def log(self, sheet_name, row_number, record_identifier, record_name, status, message):
        if not self.history:
            return
        self.pending_lines.append({
            "history_id": self.history.id,
            "sheet_name": sheet_name,
            "row_number": row_number,
            "record_identifier": str(record_identifier or ""),
            "record_name": str(record_name or ""),
            "status": status,
            "message": str(message or ""),
        })
        if len(self.pending_lines) >= 100:
            self.flush()

    def flush(self):
        if self.history and self.pending_lines:
            self.history.env["sf.import.history.line"].create(self.pending_lines)
            self.pending_lines = []


def _sheet_rows(wb, sheet, columns):
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
    # A column entirely absent from the header (not just blank cells) reads
    # as None for every row - same tolerance _move_rows already gives an
    # optional column, so a sheet can omit one it doesn't need at all.
    idx = {col: header.index(col) for col in columns if col in header}
    for row_num, row in enumerate(rows, start=2):
        values = {col: (row[idx[col]] if col in idx and idx[col] < len(row) else None) for col in columns}
        values["_row_number"] = row_num
        if values.get("id"):
            values["id"] = values["id"] if "." in str(values["id"]) else f"{MODULE}.{values['id']}"
            yield values
        else:
            if any(v is not None for k, v in values.items() if k != "_row_number"):
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
        data.write({"model": model, "res_id": record.id, "noupdate": True})
    else:
        env["ir.model.data"].create({"module": module, "name": name, "model": model, "res_id": record.id, "noupdate": True})
    return record


def _account_by_code_name(env, code_name):
    if not code_name:
        return env["account.account"]
    code = code_name.split(" ", 1)[0]
    return env["account.account"].with_context(active_test=False).search([("code", "=", code)], limit=1)


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
