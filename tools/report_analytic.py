"""Sheet account.analytic.plan / a.an.p, account.analytic.account / a.an.a,
account.report / a.r, dan account.report.line / a.r.l.
"""

import logging

from .common import MODULE, find_sheet, _parse_bool, _sheet_rows, _get_or_create

_logger = logging.getLogger(__name__)


ANALYTIC_PLAN_APPLICABILITIES = {
    "optional": "optional",
    "opsional": "optional",
    "mandatory": "mandatory",
    "wajib": "mandatory",
    "unavailable": "unavailable",
    "tidak tersedia": "unavailable",
}


REPORT_RELATED_REFS = {
    "account_reports.balance_sheet": {
        "action": "account_reports.action_account_report_bs",
        "menu": "account_reports.menu_action_account_report_balance_sheet",
    },
    "account_reports.profit_and_loss": {
        "action": "account_reports.action_account_report_pl",
        "menu": "account_reports.menu_action_account_report_profit_and_loss",
    },
    "account_reports.cash_flow_report": {
        "action": "account_reports.action_account_report_cs",
        "menu": "account_reports.menu_action_account_report_cash_flow",
    },
    "account_reports.executive_summary": {
        "action": "account_reports.action_account_report_exec_summary",
        "menu": "account_reports.menu_action_account_report_exec_summary",
    },
    "account_reports.trial_balance_report": {
        "action": "account_reports.action_account_report_coa",
        "menu": "account_reports.menu_action_account_report_coa",
    },
    "account_reports.general_ledger_report": {
        "action": "account_reports.action_account_report_general_ledger",
        "menu": "account_reports.menu_action_account_report_general_ledger",
    },
    "account_reports.partner_ledger_report": {
        "action": "account_reports.action_account_report_partner_ledger",
        "menu": "account_reports.menu_action_account_report_partner_ledger",
    },
    "account_reports.aged_receivable_report": {
        "action": "account_reports.action_account_report_ar",
        "menu": "account_reports.menu_action_account_report_aged_receivable",
    },
    "account_reports.aged_payable_report": {
        "action": "account_reports.action_account_report_ap",
        "menu": "account_reports.menu_action_account_report_aged_payable",
    },
}


def _import_account_analytic_plan(env, wb, logger=None):
    sheet = find_sheet(wb, "account.analytic.plan")
    if not sheet:
        return
    columns = ("id", "name", "sequence", "default_applicability", "color", "description", "parent_id")
    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
        values = {
            "name": row["name"],
        }
        if row["sequence"] is not None:
            values["sequence"] = int(row["sequence"])
        if row["color"] is not None:
            values["color"] = int(row["color"])
        if row["default_applicability"]:
            raw_app = str(row["default_applicability"]).strip().lower()
            if raw_app in ANALYTIC_PLAN_APPLICABILITIES:
                values["default_applicability"] = ANALYTIC_PLAN_APPLICABILITIES[raw_app]
        if row["description"]:
            values["description"] = str(row["description"])
        if row["parent_id"]:
            parent = env.ref(row["parent_id"], raise_if_not_found=False)
            if not parent and "." not in str(row["parent_id"]):
                parent = env.ref(f"{MODULE}.{row['parent_id']}", raise_if_not_found=False)
            if not parent:
                parent = env["account.analytic.plan"].search([("name", "=", row["parent_id"])], limit=1)
            if parent:
                values["parent_id"] = parent.id
        _get_or_create(env, "account.analytic.plan", row["id"], values)
        if logger:
            logger.log(
                sheet,
                row_num,
                str(row.get("id", "")),
                str(row.get("name", "")),
                "success",
                f"Rencana Analitik '{row['name']}' berhasil disinkronkan.",
            )
    if logger:
        logger.flush()


def _import_account_analytic_account(env, wb, logger=None):
    sheet = find_sheet(wb, "account.analytic.account")
    if not sheet:
        return
    columns = ("id", "name", "plan_id", "code", "active", "partner_id")
    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
        plan = None
        if row["plan_id"]:
            plan_str = str(row["plan_id"]).strip()
            plan = env.ref(plan_str, raise_if_not_found=False)
            if not plan and "." not in plan_str:
                plan = env.ref(f"{MODULE}.{plan_str}", raise_if_not_found=False)
            if not plan:
                plan = env["account.analytic.plan"].search([("name", "=", plan_str)], limit=1)
            if not plan:
                plan = env["account.analytic.plan"].search([("complete_name", "=", plan_str)], limit=1)

        if not plan:
            _logger.warning(
                "Rencana Analitik (plan_id) %r tidak ditemukan untuk akun analitik %r (id=%s) - dilewati.",
                row["plan_id"], row["name"], row["id"]
            )
            if logger:
                logger.log(
                    sheet,
                    row_num,
                    str(row.get("code") or row.get("id", "") or ""),
                    str(row.get("name", "")),
                    "warning",
                    f"Rencana Analitik (plan_id) '{row['plan_id']}' tidak ditemukan - dilewati.",
                )
            continue

        values = {
            "name": row["name"],
            "plan_id": plan.id,
        }
        if row["code"]:
            code_val = str(int(row["code"])) if isinstance(row["code"], float) and row["code"].is_integer() else str(row["code"])
            values["code"] = code_val
        if row["active"] is not None:
            parsed_active = _parse_bool(row["active"])
            if parsed_active is not None:
                values["active"] = parsed_active
        if row["partner_id"]:
            partner = env["res.partner"].search([("name", "=", row["partner_id"])], limit=1)
            if partner:
                values["partner_id"] = partner.id
        _get_or_create(env, "account.analytic.account", row["id"], values)
        if logger:
            logger.log(
                sheet,
                row_num,
                str(row.get("code") or row.get("id", "") or ""),
                str(row.get("name", "")),
                "success",
                f"Akun Analitik [{row.get('code') or ''}] {row['name']} berhasil disinkronkan.",
            )
    if logger:
        logger.flush()


def _import_account_report(env, wb, logger=None):
    sheet = find_sheet(wb, "account.report")
    if not sheet or "account.report" not in env:
        return
    columns = ("id", "name")
    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
        if not row["name"]:
            continue
        raw_id = str(row["id"]).strip()
        xml_id = raw_id
        if xml_id.startswith(f"{MODULE}.account_reports."):
            xml_id = xml_id[len(f"{MODULE}."):]
        elif xml_id.startswith(f"{MODULE}.account."):
            xml_id = xml_id[len(f"{MODULE}."):]

        report = env.ref(xml_id, raise_if_not_found=False)
        if not report and "." not in xml_id:
            report = env.ref(f"account_reports.{xml_id}", raise_if_not_found=False)
        if not report:
            report = env["account.report"].search([("name", "=", xml_id)], limit=1)

        if not report:
            _logger.warning("Laporan accounting (account.report) %r tidak ditemukan - dilewati.", xml_id)
            if logger:
                logger.log(
                    sheet,
                    row_num,
                    xml_id,
                    str(row.get("name", "")),
                    "warning",
                    f"Laporan accounting '{xml_id}' tidak ditemukan di sistem - dilewati.",
                )
            continue

        report.write({"name": row["name"]})

        lookup_id = xml_id if "." in xml_id else f"account_reports.{xml_id}"
        refs = REPORT_RELATED_REFS.get(lookup_id) or REPORT_RELATED_REFS.get(xml_id)
        if refs:
            action = env.ref(refs["action"], raise_if_not_found=False)
            if action:
                action.write({"name": row["name"]})
            menu = env.ref(refs["menu"], raise_if_not_found=False)
            if menu:
                menu.write({"name": row["name"]})
        else:
            actions = env["ir.actions.client"].search([("tag", "=", "account_report")])
            for act in actions:
                if str(report.id) in str(act.context):
                    act.write({"name": row["name"]})
                    menus = env["ir.ui.menu"].search([("action", "=", f"ir.actions.client,{act.id}")])
                    if menus:
                        menus.write({"name": row["name"]})

        if logger:
            logger.log(
                sheet,
                row_num,
                xml_id,
                str(row.get("name", "")),
                "success",
                f"Judul Laporan Keuangan '{row['name']}' berhasil disinkronkan.",
            )
    if logger:
        logger.flush()


def _import_account_report_line(env, wb, logger=None):
    sheet = find_sheet(wb, "account.report.line")
    if not sheet or "account.report.line" not in env:
        return
    columns = ("id", "name", "code")
    for row in _sheet_rows(wb, sheet, columns):
        row_num = row.get("_row_number", 0)
        if not row["name"]:
            continue
        raw_id = str(row["id"]).strip()
        xml_id = raw_id
        if xml_id.startswith(f"{MODULE}.account_reports."):
            xml_id = xml_id[len(f"{MODULE}."):]
        elif xml_id.startswith(f"{MODULE}.account."):
            xml_id = xml_id[len(f"{MODULE}."):]

        line = env.ref(xml_id, raise_if_not_found=False)
        if not line and "." not in xml_id:
            line = env.ref(f"account_reports.{xml_id}", raise_if_not_found=False)
        if not line and row.get("code"):
            line = env["account.report.line"].search([("code", "=", str(row["code"]).strip())], limit=1)
        if not line:
            line = env["account.report.line"].search([("name", "=", xml_id)], limit=1)

        if not line:
            _logger.warning("Baris laporan (account.report.line) %r tidak ditemukan - dilewati.", xml_id)
            if logger:
                logger.log(
                    sheet,
                    row_num,
                    str(row.get("code") or xml_id),
                    str(row.get("name", "")),
                    "warning",
                    f"Baris laporan '{xml_id}' tidak ditemukan di sistem - dilewati.",
                )
            continue

        line.write({"name": row["name"]})
        if logger:
            logger.log(
                sheet,
                row_num,
                str(row.get("code") or xml_id),
                str(row.get("name", "")),
                "success",
                f"Baris Laporan Keuangan '{row['name']}' berhasil disinkronkan.",
            )
    if logger:
        logger.flush()
