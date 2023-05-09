# -*- coding: utf-8 -*-
"""
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
"""
from odoo import models, fields


class KeContractType(models.Model):
    """ inherited types of contracts to add more """
    _name = "hr.contract.type"
    _inherit = "hr.contract.type"

    rem_type = fields.Selection(
        [
            ('monthly',
             'Monthly Pay'),
            ('hourly',
             'Hourly Pay'),
            ('daily',
             'Daily Pay')],
        'Remuneration Model',
        required=True,
        default='monthly',
        help="""You specify which remuneration model that is used in this contract.\n
                                [Monthly] - Employee is paid a predetermined monthly Salary\n
                                [Hourly] - Employee is paid based on number of hours worked\n
                                [Daily] - Employee is paid based on number of days worked""")

    tax_applicable = fields.Selection(
        [
            ('paye',
             'P.A.Y.E - Pay As You Earn'),
            ('wht',
             'Withholding Tax')],
        'Applicable Tax',
        required=True,
        default='paye',
        help="""Select the applicable Taxation rate based on the type of contract given \n
        to the employee [PAYE] - Applies to all employments but it does not include \
                earnings from 'casual\n
        employment' which means any engagement with any one employer which is made for a period\n
        of less than one month, the emoluments of which are calculated by reference to the\n
        period of the engagement or shorter intervals\n
        [Withholding Tax] - Applies to both resident and non resident individuals hired on\n
        consultancy agreements or terms and is deducted from consultancy fees or contractual \n
        fees and paid to KRA on monthly basis.""")


class KeContract(models.Model):
    """Inherited to add more features"""
    _inherit = "hr.contract"

    wage = fields.Monetary(
        'Remuneration',
        required=True,
        currency_field='currency_id',
        help="""This is the amount of salary or wage paid to the employee.\n
        If the remunation model is [Monthly] then this is the Monthly Basic Salary paid to
        employee\n If the remuneration model is [Daily] then this is the Daily Wage
        paid to employee. \n If the remuneration model is [Hourly] then this is the
        Hourly Wage paid to employee""")
    rem_type = fields.Selection(related="type_id.rem_type", readonly=True)
    tax_applicable = fields.Selection(
        related="type_id.tax_applicable", readonly=True)
    struct_id = fields.Many2one(
        'hr.payroll.structure',
        'Salary Structure',
        help="""Here you choose the right salary structure to use to determine the \
                pay for employee with this contract. Choose the structure with the\
                same remuneration type as the contract""")
    currency_id = fields.Many2one(
        'res.currency',
        'Currency',
        related='employee_id.currency_id')
