# -*- coding: utf-8 -*-
{
    'name': "MRP manufacturing order subcontract operation (work order)",

    'summary': """
        Allow to subcontract the operation (work order) by setting the workcenter at vendor's site.
        Purchase order to the vendor will be created with the product (service) named as "operation_name@manufacturing_order".
        
        """,

    'description': """
        Allow to subcontract the operation (work order)
    """,

    'author': "HZY",
    'website': "https://github.com/hastelloy/mrp_operation_subcontracting/",

    "license": "LGPL-3",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/16.0/odoo/addons/base/data/ir_module_category_data.xml
    # for the full list
    'category': 'Manufacturing',
    'version': '0.1',

    # any module necessary for this one to work correctly
    'depends': ['base', 'mrp', 'purchase'],

    # always loaded
    'data': [
        # 'security/ir.model.access.csv',
        # 'views/views.xml',
        'views/mrp_workcenter_views.xml',
        'views/mrp_workorder_views.xml',
        'views/mrp_production_views.xml',
        'views/templates.xml',
    ],
    # only loaded in demonstration mode
    'demo': [
        'demo/demo.xml',
    ],
}
