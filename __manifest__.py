# -*- coding: utf-8 -*-
{
    'name': "Custom Sale Report",

    'summary': "Custom Sale Report",

    'description': """
Custom Excel Sale Report
    """,
    'author': "Ayman Gamal",
    'version': '17.0.0.1',
    'depends': ['sale_management', 'account'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/xlsx_sales_wizard.xml',
        'views/menu.xml',
    ],
    # 'images': ['static/description/banner.png'],
    'license': 'AGPL-3',
    'installable': True,
    'auto_install': False,
    'application': False,
}

