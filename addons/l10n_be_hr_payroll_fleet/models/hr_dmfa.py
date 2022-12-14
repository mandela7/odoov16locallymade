# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.
from dateutil.relativedelta import relativedelta
from datetime import date
from lxml import etree

from odoo import api, fields, models, _
from odoo.tools import date_utils
from odoo.exceptions import UserError
from odoo.addons.l10n_be_hr_payroll.models.hr_dmfa import DMFANode, format_amount


class DMFACompanyVehicle(DMFANode):

    def __init__(self, vehicle, sequence=1):
        super().__init__(vehicle.env, sequence=sequence)
        self.license_plate = vehicle.license_plate
        self.eco_vehicle = -1


class HrDMFAReport(models.Model):
    _inherit = 'l10n_be.dmfa'

    vehicle_ids = fields.One2many('fleet.vehicle', compute='_compute_vehicle_ids')

    @api.depends('quarter_end')
    def _compute_vehicle_ids(self):
        monthly_pay = self.env.ref('l10n_be_hr_payroll.hr_payroll_structure_cp200_employee_salary')
        for dmfa in self:
            vehicles = self.env['hr.payslip'].search([
                # ('employee_id', 'in', employees.ids),
                ('date_to', '>=', self.quarter_start),
                ('date_to', '<=', self.quarter_end),
                ('state', 'in', ['done', 'paid']),
                ('company_id', '=', self.company_id.id),
                ('struct_id', '=', monthly_pay.id),
            ]).mapped('vehicle_id')
            dmfa.vehicle_ids = [(6, False, vehicles.ids)]

    def _get_rendering_data(self):
        invalid_vehicles = self.vehicle_ids.filtered(lambda v: len(v.license_plate) > 10)
        if invalid_vehicles:
            raise UserError(_('The following license plates are invalid:\n%s', '\n'.join(invalid_vehicles.mapped('license_plate'))))

        return dict(
            super()._get_rendering_data(),
            vehicles_cotisation=format_amount(self._get_vehicles_contribution()),
            vehicles=DMFACompanyVehicle.init_multi([(vehicle,) for vehicle in self.vehicle_ids]),
        )

    def _get_vehicles_contribution(self):
        regular_payslip = self.env.ref('l10n_be_hr_payroll.hr_payroll_structure_cp200_employee_salary')
        payslips_sudo = self.env['hr.payslip'].sudo().search([
            ('date_to', '>=', self.quarter_start),
            ('date_to', '<=', self.quarter_end),
            ('state', 'in', ['done', 'paid']),
            ('struct_id', '=', regular_payslip.id),
            ('company_id', '=', self.company_id.id),
        ])
        co2_fees = payslips_sudo._get_line_values(['CO2FEE'], compute_sum=True)['CO2FEE']['sum']['total']
        return round(co2_fees, 2)

    def _get_global_contribution(self, employees_infos, double_onss):
        # En DMFA et en DMFAPPL, la cotisation de solidarit?? sur l???usage personnel d???un v??hicule de
        # soci??t?? se d??clare globalement par cat??gorie d'employeur dans le bloc 90002 ?? cotisation
        # non li??e ?? une personne physique?? sous le code travailleur 862.
        # NB : Il est autoris?? de rassembler les donn??es de toute l???entreprise sous une seule
        # cat??gorie.
        # De plus, dans le bloc fonctionnel 90294 ?? V??hicule de soci??t?? ??, la mention des num??ros de
        # plaque des v??hicules concern??s est obligatoire. Un m??me num??ro d???immatriculation ne peut
        # ??tre repris qu???une seule fois.
        # L'avantage per??u par le travailleur pour l'usage d'un v??hicule de soci??t?? doit ??galement
        # ??tre d??clar??  sous  le code r??mun??ration DMFA 10  ou le code r??mun??ration DMFAPPL 770 dans
        # le bloc fonctionnel 90019 "R??mun??ration de l'occupation ligne travailleur".
        # Lorsque la DMFA ou la DMFAPPL  est introduite via le web, le montant global de cette
        # cotisation doit ??tre mentionn?? dans les cotisations dues pour l???ensemble de l???entreprise,
        # les num??ros de plaques des v??hicules concern??s introduits dans l?????cran pr??vu et
        # l'avantage d??clar?? avec les r??mun??rations du travailleur.
        amount = super()._get_global_contribution(employees_infos, double_onss)
        return amount + self._get_vehicles_contribution()
