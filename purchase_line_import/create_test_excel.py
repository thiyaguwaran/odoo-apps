# -*- coding: utf-8 -*-
import openpyxl

def generate_test_excels():
    # 1. Create a valid test file
    wb_valid = openpyxl.Workbook()
    ws_valid = wb_valid.active
    ws_valid.title = "Purchase Lines"

    # Headers
    headers = ["Part Number", "Quantity", "Unit Price", "Description", "UoM", "Analytic Account", "Taxes"]
    ws_valid.append(headers)

    # Data
    rows_valid = [
        ["DA02-002", 3, 150.0, "Robotic Automation Cell Type A", "Units", "3S-CM1-25133", "GST 5%"],
        ["PGN-PLUS-P 200-2-IS-SD", 5, 99.9, "Universal Gripper PGN-PLUS", "Units", "3A-BO1-25113", "GST 18%"],
        ["PGN-PLUS-P 100-1-AS-SD", 2, 80.0, "Universal Gripper PGN-PLUS-P 100", "", "", "GST 28%"]  # defaults test
    ]
    for row in rows_valid:
        ws_valid.append(row)

    wb_valid.save("test_purchase_import_valid.xlsx")
    print("Created test_purchase_import_valid.xlsx")

    # 2. Create an invalid test file (to test validation)
    wb_invalid = openpyxl.Workbook()
    ws_invalid = wb_invalid.active
    ws_invalid.title = "Purchase Lines"
    ws_invalid.append(headers)

    rows_invalid = [
        ["DA02-002", 3, 150.0, "Valid Row", "Units", "3S-CM1-25133", "GST 5%"],
        ["INVALID-PART-999", 1, 10.0, "Invalid Part Number Row", "Units", "", "GST 5%"],
        ["PGN-PLUS-P 200-2-IS-SD", "invalid_qty", 99.9, "Invalid Quantity Row", "Units", "", "GST 5%"],
        ["PGN-PLUS-P 200-2-IS-SD", 1, 99.9, "Invalid Tax Row", "Units", "", "INVALID-TAX-999"]
    ]
    for row in rows_invalid:
        ws_invalid.append(row)

    wb_invalid.save("test_purchase_import_invalid.xlsx")
    print("Created test_purchase_import_invalid.xlsx")

if __name__ == "__main__":
    generate_test_excels()
