{
    "name": "Import Data Master (xlsx)",
    "version": "19.0.1.0.0",
    "category": "SF",
    "summary": "Import CoA, jurnal, asset, kas/bank, dan opening balance dari file Excel",
    "author": "Nimdasx",
    "license": "OEEL-1",
    "depends": ["accountant"],
    "data": [
        "security/ir.model.access.csv",
        "data/ir.model.fields.selection.csv",
        "wizard/master_data_import_wizard_views.xml",
    ],
    "installable": True,
    "application": True,
    "external_dependencies": {"python": ["openpyxl"]},
}
