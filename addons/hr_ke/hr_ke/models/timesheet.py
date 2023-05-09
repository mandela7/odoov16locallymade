# -*- coding: utf-8 -*-
# License LGPL-3.0 or later (https://opensource.org/licenses/LGPL-3.0).
#
# This software and associated files (the "Software") may only be used (executed,
# modified, executed after modifications) if you have purchased a valid license
# from the authors, typically via Odoo Apps, or if you have received a written
# agreement from the authors of the Software (see the COPYRIGHT section below).
#
# You may develop Odoo modules that use the Software as a library (typically
# by depending on it, importing it and using its resources), but without copying
# any source code or material from the Software. You may distribute those
# modules under the license of your choice, provided that this license is
# compatible with the terms of the Odoo Proprietary License (For example:
# LGPL, MIT, or proprietary licenses similar to this one).
#
# It is forbidden to publish, distribute, sublicense, or sell copies of the Software
# or modified copies of the Software.
#
# The above copyright notice and this permission notice must be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
# DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
# ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.
#
#########COPYRIGHT#####
# © 2016 Bernard K Too<bernard.too@optima.co.ke>

from odoo import models, fields, api, _
from odoo.exceptions import ValidationError


class hr_ke_timesheet_sheet(models.Model):
    _inherit = 'hr_timesheet.sheet'

    slip_id = fields.Many2one(
        'hr.payslip',
        'Payslip',
        readonly=True,
        states={
            'new': [
                ('readonly',
                 False)]},
        help="This is the Payslip which the approved timesheet info will appear in")
    slip_ref = fields.Char(
        'Payslip Number',
        related='slip_id.number',
        readonly=True,
        help="This is the Payslip which the approved timesheet info will appear in")
    employee_id = fields.Many2one(
        'hr.employee', 'Employee', required=True, readonly=True, states={
            'draft': [
                ('readonly', False)], 'new': [
                ('readonly', False)]})

    # below method is called by the timesheet workflow when timesheet is being
    # approved

    def approve_timesheet(self):
        for rec in self:
            slip = {'id': ''}
            code = ''
            field_nm = ''
            arg = ''
            contract_id = rec.employee_id._get_latest_contract(field_nm, arg)
            contract = rec.env['hr.contract'].browse(
                contract_id[rec.employee_id.id])
            if contract.rem_type == 'monthly':
                code = 'PM'
            elif contract.rem_type == 'hourly':
                code = 'PH'
            else:
                code = 'PD'

            vals_payslip = {
                'employee_id': rec.employee_id.id,
                'contract_id': contract.id,
                'struct_id': contract.struct_id.id,
                'date_from': rec.date_from,
                'date_to': rec.date_to,
                'state': 'draft',
                'paid': False,
                'name': 'Payslip of ' + rec.employee_id.name_related + ' for ' + rec.date_from + ' to ' + rec.date_to,
            }
            slip = rec.env['hr.payslip'].create(vals_payslip)
            vals_work = {
                'code': code,
                'contract_id': contract.id,
                'number_of_days': rec.timesheet_activity_count,
                'number_of_hours': rec.total_timesheet,
                'name': rec.employee_id.name_related,
                'payslip_id': slip.id,
            }
            res = rec.env['hr.payslip.worked_days'].create(vals_work)
            if res:
                slip.compute_sheet()
            rec.write({'state': 'done', 'slip_id': slip.id})
            return True

    
    def action_set_to_draft(self):
        for rec in self:
            super(hr_ke_timesheet_sheet, rec).action_set_to_draft()
            # cancel payslip if still in draft state
            if rec.slip_id and rec.slip_id.state == 'draft':
                rec.slip_id.signal_workflow('cancel_sheet')
        return True
