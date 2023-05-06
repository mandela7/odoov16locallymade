# Copyright © 2018 Garazd Creation (<https://garazd.biz>)
# @author: Yurii Razumovskyi (<support@garazd.biz>)
# @author: Iryna Razumovska (<support@garazd.biz>)
# License LGPL-3.0 or later (https://www.gnu.org/licenses/lgpl-3.0.html).

from odoo import api, fields, models


class PrintProductLabelLine(models.TransientModel):
    _name = "print.product.label.line"
    _description = 'Line with a Product Label Data'

    selected = fields.Boolean(string='Print', default=True)
    wizard_id = fields.Many2one(comodel_name='print.product.label')
    product_id = fields.Many2one(comodel_name='product.product', required=True)
    barcode = fields.Char(compute='_compute_barcode')
    qty_initial = fields.Integer(string='Initial Qty', default=1)
    qty = fields.Integer(string='Label Qty', default=1)
    company_id = fields.Many2one(
        comodel_name='res.company',
        required=True,
        default=lambda self: self.env.user.company_id,
    )

    @api.depends('product_id')
    def _compute_barcode(self):
        for label in self:
            label.barcode = label.product_id.barcode

    def action_plus_qty(self):
        for record in self:
            if not record.qty:
                record.update({'selected': True})
            record.update({'qty': record.qty + 1})

    def action_minus_qty(self):
        for record in self:
            if record.qty > 0:
                record.update({'qty': record.qty - 1})
                if not record.qty:
                    record.update({'selected': False})
