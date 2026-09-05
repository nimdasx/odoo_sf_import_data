import base64
import io
import re
import requests
from markupsafe import Markup
from openpyxl import load_workbook

from odoo import api, fields, models
from odoo.exceptions import UserError
from odoo.tools import html_escape

from ..tools.import_engine import SHEET_ALIASES, run_import

GOOGLE_SHEET_ID_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")


class SfImportHistory(models.Model):
    _name = "sf.import.history"
    _description = "Riwayat Import Data Master"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "date desc, id desc"

    name = fields.Char(string="Nomor Referensi", required=True, copy=False, readonly=True, default="Draft Import")
    date = fields.Datetime(string="Tanggal Eksekusi", default=fields.Datetime.now, readonly=True, tracking=True)
    user_id = fields.Many2one(
        "res.users",
        string="User Eksekutor",
        default=lambda self: self.env.user,
        readonly=True,
        tracking=True,
    )
    company_id = fields.Many2one(
        "res.company",
        string="Perusahaan",
        default=lambda self: self.env.company,
        required=True,
        tracking=True,
    )
    source_type = fields.Selection(
        [
            ("google_sheet", "Google Sheet Link"),
            ("file", "File Excel (.xlsx)"),
        ],
        string="Sumber Data",
        default="google_sheet",
        required=True,
        tracking=True,
    )
    filename = fields.Char(string="Nama File", tracking=True)
    file = fields.Binary(string="File Excel (.xlsx)")
    google_sheet_url = fields.Char(string="Google Sheet URL", tracking=True)
    state = fields.Selection(
        [
            ("draft", "Draft"),
            ("processing", "Sedang Diproses"),
            ("completed", "Selesai"),
            ("warning", "Selesai (Ada Peringatan / Duplikat)"),
            ("failed", "Gagal"),
        ],
        string="Status",
        default="draft",
        required=True,
        tracking=True,
    )
    total_rows = fields.Integer(string="Total Baris", default=0, tracking=True)
    total_success = fields.Integer(string="Sukses", default=0, tracking=True)
    total_warning = fields.Integer(string="Peringatan / Duplikat", default=0, tracking=True)
    total_skipped = fields.Integer(string="Dilewati", default=0, tracking=True)
    total_error = fields.Integer(string="Gagal / Error", default=0, tracking=True)
    summary_html = fields.Html(string="Ringkasan Hasil")
    error_message = fields.Text(string="Pesan Error")
    line_ids = fields.One2many(
        "sf.import.history.line",
        "history_id",
        string="Semua Log per Record",
    )
    account_line_ids = fields.One2many(
        "sf.import.history.line",
        compute="_compute_sheet_lines",
        string="Log Akun (COA)",
    )
    journal_line_ids = fields.One2many(
        "sf.import.history.line",
        compute="_compute_sheet_lines",
        string="Log Jurnal",
    )
    kas_bank_line_ids = fields.One2many(
        "sf.import.history.line",
        compute="_compute_sheet_lines",
        string="Log Saldo Kas & Bank",
    )
    asset_line_ids = fields.One2many(
        "sf.import.history.line",
        compute="_compute_sheet_lines",
        string="Log Aset Tetap",
    )
    partner_line_ids = fields.One2many(
        "sf.import.history.line",
        compute="_compute_sheet_lines",
        string="Log Kontak / Partner",
    )
    move_line_ids = fields.One2many(
        "sf.import.history.line",
        compute="_compute_sheet_lines",
        string="Log Saldo Awal Hutang / Piutang",
    )
    report_analytic_line_ids = fields.One2many(
        "sf.import.history.line",
        compute="_compute_sheet_lines",
        string="Log Laporan & Analitik",
    )

    @api.depends("line_ids.sheet_name")
    def _compute_sheet_lines(self):
        account_names = set(SHEET_ALIASES.get("account.account", ())) | {"account.account", "a.a"}
        journal_names = set(SHEET_ALIASES.get("account.journal", ())) | {"account.journal", "a.j"}
        asset_names = set(SHEET_ALIASES.get("account.asset", ())) | {"account.asset", "a.as"}
        partner_names = set(SHEET_ALIASES.get("res.partner", ())) | {"res.partner", "r.p"}
        move_names = (
            set(SHEET_ALIASES.get("vendor_bill", ()))
            | set(SHEET_ALIASES.get("customer_invoice", ()))
            | {"vendor_bill", "customer_invoice", "account.move", "v.b", "c.i"}
        )
        report_analytic_names = (
            set(SHEET_ALIASES.get("company", ()))
            | set(SHEET_ALIASES.get("account.analytic.plan", ()))
            | set(SHEET_ALIASES.get("account.analytic.account", ()))
            | set(SHEET_ALIASES.get("account.report", ()))
            | set(SHEET_ALIASES.get("account.report.line", ()))
            | {"company", "c", "a.r", "a.r.l", "a.an.p", "a.an.a"}
        )

        for rec in self:
            rec.account_line_ids = rec.line_ids.filtered(lambda l: l.sheet_name in account_names)
            rec.journal_line_ids = rec.line_ids.filtered(lambda l: l.sheet_name in journal_names)
            rec.kas_bank_line_ids = rec.line_ids.filtered(
                lambda l: "Saldo Awal" in (l.sheet_name or "")
            )
            rec.asset_line_ids = rec.line_ids.filtered(lambda l: l.sheet_name in asset_names)
            rec.partner_line_ids = rec.line_ids.filtered(lambda l: l.sheet_name in partner_names)
            rec.move_line_ids = rec.line_ids.filtered(
                lambda l: l.sheet_name in move_names
                or "bill" in (l.sheet_name or "").lower()
                or "invoice" in (l.sheet_name or "").lower()
            )
            rec.report_analytic_line_ids = rec.line_ids.filtered(
                lambda l: l.sheet_name in report_analytic_names
                or "report" in (l.sheet_name or "").lower()
                or "analytic" in (l.sheet_name or "").lower()
            )

    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if not vals.get("name") or vals["name"] == "Draft Import":
                vals["name"] = self.env["ir.sequence"].next_by_code("sf.import.history") or (
                    f"IMP/{fields.Datetime.now().strftime('%Y%m%d/%H%M%S')}"
                )
        return super().create(vals_list)

    def action_run_import(self):
        self.ensure_one()
        if self.source_type == "file":
            if not self.file:
                raise UserError("Silakan upload file Excel (.xlsx) terlebih dahulu.")
            content = base64.b64decode(self.file)
        else:
            if not self.google_sheet_url:
                raise UserError("Silakan masukkan link Google Sheet terlebih dahulu.")
            content = self._download_google_sheet(self.google_sheet_url)

        try:
            wb = load_workbook(io.BytesIO(content), data_only=True)
        except Exception as e:
            raise UserError(f"File yang diimport bukan file Excel (.xlsx) yang valid: {e}") from e

        # Reset baris log lama jika re-run
        self.line_ids.unlink()
        self.write({
            "state": "processing",
            "date": fields.Datetime.now(),
            "error_message": False,
            "total_rows": 0,
            "total_success": 0,
            "total_warning": 0,
            "total_skipped": 0,
            "total_error": 0,
        })

        source_info = html_escape(self.google_sheet_url if self.source_type == "google_sheet" else (self.filename or ""))
        self.message_post(
            body=Markup(f"<b>Proses import data master dimulai</b> dari sumber: <i>{source_info}</i>"),
            message_type="notification",
            subtype_xmlid="mail.mt_note",
        )

        try:
            run_import(self.env, wb, history=self)
        except Exception as e:
            self.write({
                "state": "failed",
                "error_message": str(e),
            })
            self.message_post(
                body=Markup(f"❌ <b>Import Gagal:</b><br/>{html_escape(str(e))}"),
                message_type="comment",
                subtype_xmlid="mail.mt_comment",
            )
            raise

        # Hitung statistik akhir
        success_cnt = self.env["sf.import.history.line"].search_count([
            ("history_id", "=", self.id),
            ("status", "=", "success"),
        ])
        warning_cnt = self.env["sf.import.history.line"].search_count([
            ("history_id", "=", self.id),
            ("status", "=", "warning"),
        ])
        skipped_cnt = self.env["sf.import.history.line"].search_count([
            ("history_id", "=", self.id),
            ("status", "=", "skipped"),
        ])
        error_cnt = self.env["sf.import.history.line"].search_count([
            ("history_id", "=", self.id),
            ("status", "=", "error"),
        ])
        total_cnt = success_cnt + warning_cnt + skipped_cnt + error_cnt

        final_state = "warning" if (warning_cnt > 0 or error_cnt > 0) else "completed"

        summary = f"""
        <div class="alert alert-{'warning' if final_state == 'warning' else 'success'} m-0 p-2">
            <h5 class="mb-1 font-weight-bold">
                <i class="fa fa-{'exclamation-triangle' if final_state == 'warning' else 'check-circle'} mr-1"></i>
                Import Selesai {'dengan Catatan' if final_state == 'warning' else 'Sukses'}
            </h5>
            <p class="mb-0">
                Total baris diproses: <b>{total_cnt}</b> | 
                Sukses: <b class="text-success">{success_cnt}</b> | 
                Peringatan/Duplikat: <b class="text-warning">{warning_cnt}</b> | 
                Dilewati: <b class="text-muted">{skipped_cnt}</b> | 
                Error: <b class="text-danger">{error_cnt}</b>
            </p>
        </div>
        """

        self.write({
            "state": final_state,
            "total_rows": total_cnt,
            "total_success": success_cnt,
            "total_warning": warning_cnt,
            "total_skipped": skipped_cnt,
            "total_error": error_cnt,
            "summary_html": summary,
        })

        if final_state == "warning":
            msg_body = Markup(
                f"⚠️ <b>Import data master selesai dengan peringatan/duplikat</b><br/>"
                f"• Total Baris Diproses: <b>{total_cnt}</b><br/>"
                f"• Sukses: <b>{success_cnt}</b><br/>"
                f"• Peringatan: <b style='color:#e67e22;'>{warning_cnt}</b><br/>"
                f"• Dilewati: <b>{skipped_cnt}</b><br/>"
                f"• Gagal: <b style='color:#e74c3c;'>{error_cnt}</b>"
            )
        else:
            msg_body = Markup(
                f"✅ <b>Import data master selesai dengan sukses</b><br/>"
                f"• Total Baris Diproses: <b>{total_cnt}</b> (seluruhnya berhasil disinkronkan)."
            )
        self.message_post(body=msg_body, message_type="notification", subtype_xmlid="mail.mt_note")
        return True

    def action_view_lines(self):
        self.ensure_one()
        return {
            "name": f"Log Baris: {self.name}",
            "type": "ir.actions.act_window",
            "res_model": "sf.import.history.line",
            "view_mode": "list,pivot,graph,form",
            "domain": [("history_id", "=", self.id)],
            "context": {"default_history_id": self.id},
        }

    def action_view_pivot(self):
        self.ensure_one()
        return {
            "name": f"Analisis Pivot: {self.name}",
            "type": "ir.actions.act_window",
            "res_model": "sf.import.history.line",
            "view_mode": "pivot,graph,list,form",
            "domain": [("history_id", "=", self.id)],
            "context": {"default_history_id": self.id},
        }

    def action_view_warnings(self):
        self.ensure_one()
        return {
            "name": f"Peringatan / Duplikat: {self.name}",
            "type": "ir.actions.act_window",
            "res_model": "sf.import.history.line",
            "view_mode": "list,pivot,graph,form",
            "domain": [("history_id", "=", self.id), ("status", "=", "warning")],
            "context": {"default_history_id": self.id, "search_default_filter_warning": 1},
        }

    def action_view_errors(self):
        self.ensure_one()
        return {
            "name": f"Log Gagal: {self.name}",
            "type": "ir.actions.act_window",
            "res_model": "sf.import.history.line",
            "view_mode": "list,pivot,graph,form",
            "domain": [("history_id", "=", self.id), ("status", "=", "error")],
            "context": {"default_history_id": self.id, "search_default_filter_error": 1},
        }

    def _download_google_sheet(self, url):
        match = GOOGLE_SHEET_ID_RE.search(url)
        if not match:
            raise UserError(
                "Link Google Sheet tidak valid - harus berupa link spreadsheet "
                "(format: docs.google.com/spreadsheets/d/<ID>/...)."
            )
        export_url = f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=xlsx"

        try:
            response = requests.get(export_url, timeout=45)
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


class SfImportHistoryLine(models.Model):
    _name = "sf.import.history.line"
    _description = "Detail Log Baris Import"
    _order = "history_id desc, id asc"

    history_id = fields.Many2one(
        "sf.import.history",
        string="Riwayat Import",
        required=True,
        ondelete="cascade",
    )
    sheet_name = fields.Char(string="Nama Sheet", required=True)
    row_number = fields.Integer(string="Baris Excel")
    record_identifier = fields.Char(string="Kode / ID")
    record_name = fields.Char(string="Nama Record")
    status = fields.Selection(
        [
            ("success", "Sukses"),
            ("warning", "Peringatan / Duplikat"),
            ("skipped", "Dilewati"),
            ("error", "Gagal"),
        ],
        string="Status",
        default="success",
        required=True,
    )
    message = fields.Text(string="Keterangan Detail")

    @api.depends("sheet_name", "row_number", "record_identifier", "record_name")
    def _compute_display_name(self):
        for rec in self:
            rec.display_name = f"[{rec.sheet_name} Baris {rec.row_number}] {rec.record_identifier or ''} {rec.record_name or ''}".strip()
