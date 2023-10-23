from odoo import models, fields, api
import xlrd

class ExcelImportWizard(models.TransientModel):
    _name = 'excel.import.wizard'
    _description = 'Wizard pour Importer depuis Excel'

    file = fields.Binary(string='Fichier Excel', required=True)

    def import_excel_data(self):
        import base64
        import io

        file_content = base64.b64decode(self.file)
        excel_file = xlrd.open_workbook(file_contents=file_content)
        sheet = excel_file.sheet_by_index(0)

        receptions = {}
        for i in range(1, sheet.nrows):  # Ignorer la première ligne (en-têtes)
            type_reception = sheet.cell_value(i, 0)
            fournisseur_name = sheet.cell_value(i, 1)
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

            fournisseur = self.env['res.partner'].search([('name', '=', fournisseur_name), ('supplier', '=', True)], limit=1)
            if not fournisseur:
                fournisseur = self.env['res.partner'].create({
                    'name': fournisseur_name,
                    'supplier': True
                })

            if numero_commande not in receptions:
                reception = self.env['stock.picking'].create({
                    'picking_type_id': 1,
                    'partner_id': fournisseur.id,
                    'scheduled_date': date,
                    'origin': numero_commande
                })
                receptions[numero_commande] = reception
            else:
                reception = receptions[numero_commande]

            if lot_ids:
                lot = self.env['stock.production.lot'].search([('name', '=', lot_ids)], limit=1)
                if not lot:
                    lot = self.env['stock.production.lot'].create({
                        'name': lot_ids,
                        'product_id': product.id
                    })

            self.env['stock.move'].create({
                'name': nom_article,
                'product_id': product.id,
                'product_uom_qty': quantite,
                'picking_id': reception.id,
                'lot_ids': lot_ids,
            })

        return {'type': 'ir.actions.act_window_close'}
