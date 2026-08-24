import base64
import io
import re

import requests
from openpyxl import load_workbook

from odoo import fields, models
from odoo.exceptions import UserError

from ..hooks import run_import

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
        if self.file:
            if not self.filename or not self.filename.lower().endswith(".xlsx"):
                raise UserError("File harus berformat .xlsx")
            content = base64.b64decode(self.file)
        elif self.google_sheet_url:
            content = self._download_google_sheet(self.google_sheet_url)
        else:
            raise UserError("Isi salah satu: upload File Excel atau Link Google Sheet.")

        try:
            wb = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        except Exception as e:
            raise UserError(
                "File yang diimport bukan file Excel (.xlsx) yang valid."
            ) from e

        run_import(self.env, wb)

        # Push the success toast over the bus instead of returning a client
        # action, so the wizard can close itself (act_window_close) in the
        # same response instead of staying open showing the notification.
        self.env["bus.bus"]._sendone(self.env.user.partner_id, "simple_notification", {
            "title": "Import selesai",
            "message": "Data master berhasil diimport.",
            "type": "success",
            "sticky": False,
        })
        return {"type": "ir.actions.act_window_close"}

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
