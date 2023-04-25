# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import fields, models, api, Command

class MrpWorkorder(models.Model):
    _inherit = 'mrp.workorder'

    purchase_order_id = fields.Many2one("purchase.order",
        compute="_compute_purchase_order_id")

    def _calc_operation_cost(self, opt):
        duration_expected = opt.duration_expected or 0
        return (duration_expected / 60) * opt.workcenter_id.costs_hour

    def _compute_purchase_order_id(self):
        for workorder in self:
            workcenter = workorder.workcenter_id
            if not workcenter.is_subcontract:
                workorder.purchase_order_id = None
                continue
            name = f"{workorder.id}_{workorder.name}@{workorder.production_id.name}"
            # create service product
            products = self.env['product.product'].search([('name', '=', name)])
            if len(products) == 0:
                product = self.env['product.product'].create([{
                    'name': name,
                    'sale_ok': False,
                    'purchase_ok': True,
                    'detailed_type': "service"
                }])
            else:
                product = products[0]
            po = self.env['purchase.order'].search([('origin', '=', name)], limit=1)
            if not po:
                cost = self._calc_operation_cost(workorder)
                workorder.purchase_order_id = self.env['purchase.order'].create({
                    'partner_id': workcenter.partner_id.id,  
                    'state': 'to approve',
                    'origin': name,
                    'order_line': [Command.create({
                        'product_id': product.id,
                        'product_type': product.detailed_type,
                        'product_qty': workorder.production_id.product_qty,
                        'price_unit': cost
                        }),]
                }).id
            else:
                workorder.purchase_order_id = po.id                
