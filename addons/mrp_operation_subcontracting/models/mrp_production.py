from odoo import api, fields, models, _, Command

class MrpProduction(models.Model):
    """ Manufacturing Orders """
    _inherit = 'mrp.production'

    def _create_subcontract_purchase_orders(self):
        for production in self:
            for workorder in production.workorder_ids:
            

    @api.depends('bom_id', 'product_id', 'product_qty', 'product_uom_id')
    def _compute_workorder_ids(self):
        super(MrpProduction, self)._compute_workorder_ids()
        self._create_subcontract_purchase_orders()