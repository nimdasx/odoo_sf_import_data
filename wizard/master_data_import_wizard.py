import base64
import io

from openpyxl import load_workbook

from odoo import fields, models
from odoo.exceptions import UserError

from ..hooks import run_import


class MasterDataImportWizard(models.TransientModel):
    _name = "sf.import.data.wizard"
    _description = "Import Data Master (xlsx)"

    file = fields.Binary(string="File Excel (.xlsx)", required=True)
    filename = fields.Char(string="Nama File")

    def action_import(self):
        self.ensure_one()
        if not self.filename or not self.filename.lower().endswith(".xlsx"):
            raise UserError("File harus berformat .xlsx")

        content = base64.b64decode(self.file)
        wb = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
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
