import base64
import io
import re

import requests
from openpyxl import load_workbook

from odoo import fields, models
from odoo.exceptions import UserError

from ..tools.import_engine import run_import

GOOGLE_SHEET_ID_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")


class MasterDataImportWizard(models.TransientModel):
    _name = "sf.import.data.wizard"
    _description = "Import Data Master (xlsx)"

    file = fields.Binary(string="File Excel (.xlsx)")
    filename = fields.Char(string="Nama File")
    google_sheet_url = fields.Char(
        string="atau Link Google Sheet",
        help='Paste link share Google Sheets-nya (harus di-share sebagai "Anyone with '
        'the link"/public). Kalau field File di atas juga diisi, File yang dipakai.',
    )

    def action_import(self):
        self.ensure_one()
        if not self.file and not self.google_sheet_url:
            raise UserError("Isi salah satu: upload File Excel atau Link Google Sheet.")

        if self.file and (not self.filename or not self.filename.lower().endswith(".xlsx")):
            raise UserError("File harus berformat .xlsx")

        history_vals = {
            "source_type": "file" if self.file else "google_sheet",
            "filename": self.filename,
            "file": self.file,
            "google_sheet_url": self.google_sheet_url,
            "company_id": self.env.company.id,
            "user_id": self.env.user.id,
        }
        history = self.env["sf.import.history"].create(history_vals)
        history.action_run_import()

        return {
            "name": f"Riwayat Import: {history.name}",
            "type": "ir.actions.act_window",
            "res_model": "sf.import.history",
            "res_id": history.id,
            "view_mode": "form",
            "target": "current",
        }

    def _download_google_sheet(self, url):
        """Convert a Google Sheets share link (…/edit?usp=sharing, …/edit#gid=0,
        etc.) into its .xlsx export link and download it. Only works for a
        sheet shared as "Anyone with the link" - a private one 302-redirects
        to a Google sign-in page instead of the file, which downloads fine
        as far as requests is concerned but isn't a real xlsx, so that's
        caught here with a message pointing at the actual cause instead of
        surfacing as an opaque "not a valid Excel file" error later.
        """
        match = GOOGLE_SHEET_ID_RE.search(url)
        if not match:
            raise UserError(
                "Link Google Sheet tidak valid - harus berupa link spreadsheet "
                "(format: docs.google.com/spreadsheets/d/<ID>/...)."
            )
        export_url = f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=xlsx"

        try:
            response = requests.get(export_url, timeout=30)
            response.raise_for_status()
        except requests.RequestException as e:
            raise UserError(f"Gagal download dari Google Sheet: {e}") from e

        content_type = response.headers.get("Content-Type", "")
        if "spreadsheet" not in content_type:
            raise UserError(
                "Gagal download file dari Google Sheet - pastikan link sudah di-share "
                'sebagai "Anyone with the link" (public), bukan private/restricted.'
            )
        return response.content
