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
        product = False
        for i in range(1, sheet.nrows):  # Ignorer la première ligne (en-têtes)
            type_reception = sheet.cell_value(i, 0)
            fournisseur_name = sheet.cell_value(i, 1)
            date = sheet.cell_value(i, 2)
            numero_commande = sheet.cell_value(i, 3)
            reference_article = sheet.cell_value(i, 4)
            nom_article = sheet.cell_value(i, 5)
            quantite = sheet.cell_value(i, 6)
            lot_ids = sheet.cell_value(i, 7)
            lot = False

            product = self.env['product.product'].search([('default_code', '=', reference_article)], limit=1)
            if not product:
                product = self.env['product.product'].create({
                    'default_code': reference_article,
                    'name': nom_article
                })

            fournisseur = self.env['res.partner'].search([('name', '=', fournisseur_name)], limit=1)
            if not fournisseur:
                fournisseur = self.env['res.partner'].create({
                    'name': fournisseur_name,
                    
                })
            
            if numero_commande not in receptions:
                reception = self.env['stock.picking'].create({
                    'picking_type_id': 1,
                    'partner_id': fournisseur.id,
                    # 'scheduled_date': date,
                    'origin': numero_commande
                })
                receptions[numero_commande] = reception
            else:
                reception = receptions[numero_commande]
            lot = False
            if lot_ids and lot_ids != "":
                lot = self.env['stock.lot'].search([('name', '=', lot_ids)], limit=1)
                if not lot:
                    lot = self.env['stock.lot'].create({
                        'name': lot_ids,
                        'product_id': product.id
                    })
            list_product = []
            for pp in reception.move_ids_without_package:
                list_product.append(pp.product_id.id)
            if lot and product.id not in list_product:
             move = self.env['stock.move'].create({
                'name': nom_article,
                'product_id': product.id,
                'product_uom_qty': quantite,
                'picking_id': reception.id,
                'location_id': 5,
                'location_dest_id':8 , 
                'lot_ids': [(4,lot.id)],
            })
            elif not lot and product.id not in list_product:
             move = self.env['stock.move'].create({
                'name': nom_article,
                'product_id': product.id,
                'product_uom_qty': quantite,
                'picking_id': reception.id,
                'location_id': 5,
                'location_dest_id':8 , 
            })
            if move:
             if lot :
                moveline = self.env['stock.move.line'].create({
                'name': nom_article,
                'move_id': move.id,
                'product_id': product.id,
                'product_uom_qty': 1,
                'location_id': 5,
                'location_dest_id':8 , 
                'lot_ids': lot.id, 
                     })
             else :
                moveline = self.env['stock.move.line'].create({
                'name': nom_article,
                'move_id': move.id,
                'product_id': product.id,
                'product_uom_qty': 1,
                'location_id': 5,
                'location_dest_id':8 , 
 
                     })
            else: 
                themove = False
                for mv in reception.move_ids_without_package:
                    if mv.product_id.id == product.id:
                        themove = mv.id
                if lot:
                 moveline = self.env['stock.move.line'].create({
                 'name': nom_article,
                 'move_id': themove,
                 'product_id': product.id,
                 'product_uom_qty': 1,
                 'location_id': 5,
                 'location_dest_id':8 , 
                 'lot_id': lot.id, 
                   })
                else:
                 moveline = self.env['stock.move.line'].create({
                 'name': nom_article,
                 'move_id': themove,
                 'product_id': product.id,
                 'product_uom_qty': 1,
                 'location_id': 5,
                 'location_dest_id':8 , 
                 
                   })
        return {'type': 'ir.actions.act_window_close'}
