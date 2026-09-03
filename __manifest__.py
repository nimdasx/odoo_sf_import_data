{
    "name": "Import Data Master",
    "version": "19.0.1.1.0",
    "category": "SF",
    "summary": "Import CoA, jurnal, asset, kas/bank, dan opening balance dari file Excel dengan riwayat dan logging",
    "author": "Nimdasx",
    "license": "OEEL-1",
    "depends": ["accountant"],
    "data": [
        "security/ir.model.access.csv",
        "data/ir.model.fields.selection.csv",
        "data/sf_import_history_sequence.xml",
        "views/import_history_views.xml",
        "wizard/master_data_import_wizard_views.xml",
    ],
    "installable": True,
    "application": True,
    "external_dependencies": {"python": ["openpyxl", "requests"]},
}
