# -*- encoding: utf-8 -*-

from odoo import models,fields, api,tools
from odoo.exceptions import ValidationError

class ExcelImportWizard(models.TransientModel):
    _name = 'excel.import.wizard'
    _description = 'Wizard pour Importer depuis Excel'

    file = fields.Binary(string='Fichier Excel', required=True)

    def import_excel_data(self):
        import xlrd
        import base64
        from io import BytesIO

        excel_file = xlrd.open_workbook(file_contents=BytesIO(base64.b64decode(self.file)))
        sheet = excel_file.sheet_by_index(0)

        receptions = {}
        for i in range(1, sheet.nrows):  # Ignorer la première ligne (en-têtes)
            type_reception = sheet.cell_value(i, 0)
            fournisseur = sheet.cell_value(i, 1)
            date = sheet.cell_value(i, 2)
            numero_commande = sheet.cell_value(i, 3)
            reference_article = sheet.cell_value(i, 4)
            nom_article = sheet.cell_value(i, 5)
            quantite = sheet.cell_value(i, 6)
            lot_ids = sheet.cell_value(i, 7)

            product = self.env['product.product'].search([('default_code', '=', reference_article)], limit=1)
            if not product:
                product = self.env['product.product'].create({
                    'default_code': reference_article,
                    'name': nom_article
                })

            if numero_commande not in receptions:
                reception = self.env['stock.picking'].create({
                    'picking_type_id': type_reception,
                    'partner_id': fournisseur,
                    'scheduled_date': date,
                    'origin': numero_commande
                })
                receptions[numero_commande] = reception
            else:
                reception = receptions[numero_commande]

            self.env['stock.move'].create({
                'name': nom_article,
                'product_id': product.id,
                'product_uom_qty': quantite,
                'picking_id': reception.id,
                'lot_ids': lot_ids,
            })

        return {'type': 'ir.actions.act_window_close'}

