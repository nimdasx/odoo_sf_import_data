def migrate(cr, version):
    """Ensure dynamic business records created by import tools have noupdate=true
    so that Odoo's module upgrade does not attempt to purge/unlink them as orphaned
    module records (which violates foreign key constraints on account_move_line, etc.).
    """
    cr.execute("""
        UPDATE ir_model_data
        SET noupdate = true
        WHERE module = 'odoo_sf_import_data'
          AND model NOT IN (
              'ir.actions.act_window',
              'ir.model',
              'ir.model.access',
              'ir.model.fields',
              'ir.ui.menu',
              'ir.ui.view'
          );
    """)
