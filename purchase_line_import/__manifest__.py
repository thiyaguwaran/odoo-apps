# -*- coding: utf-8 -*-
{
    'name': 'Purchase Line Excel Import',
    'version': '18.0.1.0.0',
    'summary': 'Import Purchase Order Lines from Excel files',
    'description': 'Allows importing purchase order lines from Excel (.xlsx) files.',
    'category': 'Purchase',
    'author': 'Thiyagesh S',
    'depends': ['purchase'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/purchase_line_import_wizard.xml',
        'views/purchase_order_views.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
    'license': 'OPL-1',
    'price': 4.99,
    'currency': 'USD',
}
