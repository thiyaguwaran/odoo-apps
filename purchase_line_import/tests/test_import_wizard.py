# -*- coding: utf-8 -*-
import io
import base64
import openpyxl
from odoo.tests import common
from odoo.exceptions import ValidationError

class TestPurchaseOrderLineImportWizard(common.TransactionCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        # Set up a vendor
        cls.vendor = cls.env['res.partner'].create({
            'name': 'Test Vendor',
        })
        
        # Set up products
        cls.product_1 = cls.env['product.product'].create({
            'name': 'Product 1',
            'default_code': 'PROD-001',
            'list_price': 100.0,
            'purchase_ok': True,
        })
        cls.product_2 = cls.env['product.product'].create({
            'name': 'Product 2',
            'default_code': 'PROD-002',
            'list_price': 200.0,
            'purchase_ok': True,
        })
        
        # Set up UoM (using default Units)
        cls.uom_unit = cls.env.ref('uom.product_uom_unit')
        
        # Set up analytic account
        cls.analytic_account = cls.env['account.analytic.account'].create({
            'name': 'Test Analytic',
            'plan_id': cls.env['account.analytic.plan'].search([], limit=1).id,
        })
        
        # Set up tax
        cls.tax = cls.env['account.tax'].create({
            'name': 'Test Purchase Tax',
            'type_tax_use': 'purchase',
            'amount': 10.0,
            'company_id': cls.env.company.id,
        })
        
        # Create draft purchase order
        cls.purchase_order = cls.env['purchase.order'].create({
            'partner_id': cls.vendor.id,
        })

    def _generate_excel_base64(self, rows):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Lines"
        for row in rows:
            ws.append(row)
        
        f = io.BytesIO()
        wb.save(f)
        return base64.b64encode(f.getvalue())

    def test_valid_import(self):
        # Prepare valid rows
        # Headers: Part Number, Quantity, Unit Price, Description, UoM, Analytic Account, Taxes
        rows = [
            ["Part Number", "Quantity", "Unit Price", "Description", "UoM", "Analytic Account", "Taxes"],
            ["PROD-001", 3.0, 150.0, "Custom Prod 1", "Units", self.analytic_account.name, self.tax.name],
            ["PROD-002", 5.0, 220.0, "Custom Prod 2", "", "", ""], # defaults
        ]
        
        excel_data = self._generate_excel_base64(rows)
        
        wizard = self.env['purchase.order.line.import.wizard'].create({
            'data_file': excel_data,
            'file_name': 'test.xlsx',
        })
        
        # Run action_import
        res = wizard.with_context(active_id=self.purchase_order.id).action_import()
        self.assertEqual(res.get('type'), 'ir.actions.act_window_close')
        
        # Verify lines
        lines = self.purchase_order.order_line
        self.assertEqual(len(lines), 2)
        
        # Line 1 verification
        line1 = lines.filtered(lambda l: l.product_id == self.product_1)
        self.assertTrue(line1)
        self.assertEqual(line1.product_qty, 3.0)
        self.assertEqual(line1.price_unit, 150.0)
        self.assertEqual(line1.name, "Custom Prod 1")
        self.assertEqual(line1.product_uom, self.uom_unit)
        self.assertIn(str(self.analytic_account.id), line1.analytic_distribution)
        self.assertEqual(line1.analytic_distribution[str(self.analytic_account.id)], 100.0)
        self.assertIn(self.tax, line1.taxes_id)
        
        # Line 2 verification (uses defaults for some fields)
        line2 = lines.filtered(lambda l: l.product_id == self.product_2)
        self.assertTrue(line2)
        self.assertEqual(line2.product_qty, 5.0)
        self.assertEqual(line2.price_unit, 220.0)
        self.assertEqual(line2.name, "Custom Prod 2")
        self.assertFalse(line2.analytic_distribution)
        
    def test_invalid_import_missing_part_number(self):
        rows = [
            ["Part Number", "Quantity", "Unit Price"],
            ["", 3.0, 150.0],
        ]
        excel_data = self._generate_excel_base64(rows)
        
        wizard = self.env['purchase.order.line.import.wizard'].create({
            'data_file': excel_data,
            'file_name': 'test.xlsx',
        })
        
        with self.assertRaises(ValidationError) as error:
            wizard.with_context(active_id=self.purchase_order.id).action_import()
        self.assertIn("Row 2: Missing Part Number", error.exception.name)

    def test_invalid_import_product_not_found(self):
        rows = [
            ["Part Number", "Quantity", "Unit Price"],
            ["NON_EXISTENT_PROD", 3.0, 150.0],
        ]
        excel_data = self._generate_excel_base64(rows)
        
        wizard = self.env['purchase.order.line.import.wizard'].create({
            'data_file': excel_data,
            'file_name': 'test.xlsx',
        })
        
        with self.assertRaises(ValidationError) as error:
            wizard.with_context(active_id=self.purchase_order.id).action_import()
        self.assertIn("Row 2: Product not found", error.exception.name)
