# -*- coding: utf-8 -*-
{
    'name': 'Purchase Line Excel Import',
    'version': '18.0.1.0.0',
    'summary': 'Import purchase order lines from Excel files',
    'description': """
Import Purchase Order Lines from Excel files.

Features:
- Import PO lines from XLSX
- Product validation
- Quantity and price mapping
- Error handling
""",
    'category': 'Purchases',
    'author': 'Thiyagesh S',
    'website': 'https://github.com/thiyaguwaran/odoo-apps',
    'depends': ['purchase'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/purchase_line_import_wizard.xml',
        'views/purchase_order_views.xml',
    ],
    'license': 'OPL-1',
    'price': 4.99,
    'currency': 'USD',
    'images': ['static/description/banner.png'],
    'installable': True,
    'application': False,
    'auto_install': False,
}
