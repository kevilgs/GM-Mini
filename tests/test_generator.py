import unittest
import os
import openpyxl
import datetime
from core.generator import generate_daily_excel, EXCEL_COMMODITY_ROWS, EXCEL_DIVISION_ROWS
from core.gsheets_db import COMMODITIES, DIVISIONS

class TestGenerator(unittest.TestCase):
    def setUp(self):
        self.template_path = os.path.join("assets", "GM MINI- AUGUST 2026.xlsx")
        self.test_output_dir = "generated_workbooks"

    def _build_dummy_report_data(self, day_wagons=420.0):
        report_data = {
            'commodity': {},
            'division': {},
            'summary': {}
        }
        for c in COMMODITIES:
            report_data['commodity'][c] = {
                'rakes': 5.0 if c == 'CEMT' else 0.0,
                'wagons': day_wagons if c == 'CEMT' else 0.0,
                'cur_mtd_avg': day_wagons if c == 'CEMT' else 0.0,
                'ly_wagons': 390.0 if c == 'CEMT' else 0.0,
                'ly_mtd_avg': 390.0 if c == 'CEMT' else 0.0,
                'diff_avg': 30.0 if c == 'CEMT' else 0.0,
                'ly_whole_month_avg': 400.0 if c == 'CEMT' else 0.0,
                'diff_whole_month_avg': 20.0 if c == 'CEMT' else 0.0,
                'cur_fy_avg': day_wagons if c == 'CEMT' else 0.0,
                'ly_fy_avg': 390.0 if c == 'CEMT' else 0.0,
                'ly_full_year_avg': 410.0 if c == 'CEMT' else 0.0
            }
        for d in DIVISIONS:
            report_data['division'][d] = {
                'rakes': 4.0 if d == 'ADI' else (3.0 if d == 'GIMB' else 0.0),
                'wagons': 200.0 if d == 'ADI' else (150.0 if d == 'GIMB' else 0.0),
                'cur_mtd_avg': 200.0 if d == 'ADI' else (150.0 if d == 'GIMB' else 0.0),
                'ly_wagons': 180.0 if d == 'ADI' else 0.0,
                'ly_mtd_avg': 180.0 if d == 'ADI' else 0.0,
                'diff_avg': 20.0 if d == 'ADI' else 0.0,
                'ly_whole_month_avg': 190.0 if d == 'ADI' else 0.0,
                'diff_whole_month_avg': 10.0 if d == 'ADI' else 0.0,
                'cur_fy_avg': 200.0 if d == 'ADI' else (150.0 if d == 'GIMB' else 0.0),
                'ly_fy_avg': 180.0 if d == 'ADI' else 0.0,
                'ly_full_year_avg': 195.0 if d == 'ADI' else 0.0
            }
        report_data['summary'] = {
            'cur_today_wagons': 350.0,
            'cur_today_rakes': 7.0,
            'ly_today_wagons': 300.0,
            'cur_today_mt': 0.02,
            'ly_today_mt': 0.02,
            'cur_mtd_wagons': 350.0,
            'ly_mtd_wagons': 300.0,
            'cur_mtd_mt': 0.02,
            'ly_mtd_mt': 0.02,
            'cur_prev_months_wagons': 0.0,
            'ly_prev_months_wagons': 0.0,
            'cur_prev_months_mt': 0.0,
            'ly_prev_months_mt': 0.0,
            'cur_fy_prog_wagons': 350.0,
            'ly_fy_prog_wagons': 300.0,
            'div_cumms': {
                'ADI_GIMB': {'cur': 350.0, 'ly': 300.0},
                'BCT': {'cur': 0.0, 'ly': 0.0},
                'BRC': {'cur': 0.0, 'ly': 0.0},
                'RJT': {'cur': 0.0, 'ly': 0.0},
                'BVC': {'cur': 0.0, 'ly': 0.0},
                'RTM': {'cur': 0.0, 'ly': 0.0},
            }
        }
        return report_data

    def test_april_1st_workbook_generation(self):
        """Verify generation of April 1st report:
        - Output workbook exists
        - Single day tab named '01-04-2026'
        - LINK and Sheet1 preserved
        - Dynamic headers updated
        - Right-hand summary panel has 0 prior loading and correct today values
        """
        report_data = self._build_dummy_report_data()
        output_file = generate_daily_excel(4, 2026, 1, report_data, template_path=self.template_path)
        
        self.assertTrue(os.path.exists(output_file), f"Output file {output_file} was not created")
        
        wb = openpyxl.load_workbook(output_file)
        self.assertIn("01-04-2026", wb.sheetnames)
        self.assertNotIn("2", wb.sheetnames, "Day 2 sheet should have been pruned")
        self.assertNotIn("31", wb.sheetnames, "Day 31 sheet should have been pruned")
        self.assertIn("LINK", wb.sheetnames, "LINK sheet must be preserved for formulas")
        
        ws = wb["01-04-2026"]
        
        # 1. Header checks
        self.assertEqual(ws.cell(row=2, column=14).value, "LOADING TARGET   2026-27")
        self.assertEqual(ws.cell(row=3, column=2).value, "Target per day for  APR-26")
        self.assertIn("APR-2025", ws.cell(row=3, column=9).value)
        self.assertEqual(ws.cell(row=5, column=14).value, "Loading Target 2026-2027")
        self.assertEqual(ws.cell(row=6, column=14).value, "Progressive Loading Up To MARCH")
        
        # 2. Left panel commodity row checks (CEMT is row 5)
        cemt_row = EXCEL_COMMODITY_ROWS['CEMT']
        self.assertEqual(ws.cell(row=cemt_row, column=3).value, 5.0)   # Rakes (Col C)
        self.assertEqual(ws.cell(row=cemt_row, column=4).value, 420.0) # Wagons (Col D)
        self.assertEqual(ws.cell(row=cemt_row, column=5).value, 420.0) # Cur MTD Avg (Col E)
        self.assertEqual(ws.cell(row=cemt_row, column=6).value, 390.0) # LY Wagons (Col F)
        self.assertEqual(ws.cell(row=cemt_row, column=8).value, 30.0)  # Diff Avg (Col H)
        
        # 3. Table total formula checks
        self.assertEqual(ws.cell(row=18, column=3).value, "=SUM(C5:C17)")
        self.assertEqual(ws.cell(row=18, column=4).value, "=SUM(D5:D17)")
        self.assertEqual(ws.cell(row=28, column=3).value, "=SUM(C21:C27)")
        self.assertEqual(ws.cell(row=28, column=4).value, "=SUM(D21:D27)")
        
        # 4. Right-hand summary panel checks
        # Row 6: Progressive Loading Up To Last Month (strictly 0 for April)
        self.assertEqual(ws.cell(row=6, column=16).value, 0.0, "LY prior months wagons must be 0 for April")
        self.assertEqual(ws.cell(row=6, column=17).value, 0.0, "Cur prior months wagons must be 0 for April")
        self.assertEqual(ws.cell(row=6, column=18).value, 0.0, "LY prior months MT must be 0 for April")
        self.assertEqual(ws.cell(row=6, column=19).value, 0.0, "Cur prior months MT must be 0 for April")
        
        # Row 7: Today's loading
        self.assertEqual(ws.cell(row=7, column=17).value, 350.0)
        self.assertEqual(ws.cell(row=7, column=19).value, 0.02)
        
        # Row 8: Current month loading till date
        self.assertEqual(ws.cell(row=8, column=17).value, 350.0)
        
        # Row 10: ADI + GIMB cumulative
        self.assertEqual(ws.cell(row=10, column=17).value, 350.0)
        
        # Row 16: Progressive LDG from April
        self.assertEqual(ws.cell(row=16, column=17).value, 350.0)
        self.assertEqual(ws.cell(row=16, column=19).value, "=SUM(S10:S15)")
        
        # Row 17: To Achieve Target
        self.assertEqual(ws.cell(row=17, column=17).value, 2208615 - 350.0)
        self.assertEqual(ws.cell(row=17, column=19).value, "=+S5-S16")
        
        # 5. Hardcoded stock clearing check in P21:P27 (cleared to 0)
        for r in range(21, 28):
            self.assertEqual(ws.cell(row=r, column=16).value, 0, f"Cell P{r} should be cleared to 0")
        self.assertEqual(ws.cell(row=28, column=16).value, "=SUM(P21:P27)")
            
        # 6. Automatic calculation flags check
        self.assertTrue(wb.calculation.fullCalcOnLoad)
        self.assertTrue(wb.calculation.forceFullCalc)
        
        wb.close()

    def test_mid_year_month_header_august(self):
        """Verify month headers for August: 'Progressive Loading Up To JULY'."""
        report_data = self._build_dummy_report_data()
        output_file = generate_daily_excel(8, 2026, 15, report_data, template_path=self.template_path)
        
        wb = openpyxl.load_workbook(output_file)
        ws = wb["15-08-2026"]
        
        self.assertEqual(ws.cell(row=3, column=2).value, "Target per day for  AUG-26")
        self.assertEqual(ws.cell(row=6, column=14).value, "Progressive Loading Up To JULY")
        wb.close()

    def test_missing_template_raises_error(self):
        """Verify FileNotFoundError is raised if template path does not exist."""
        report_data = self._build_dummy_report_data()
        with self.assertRaises(FileNotFoundError):
            generate_daily_excel(4, 2026, 1, report_data, template_path="non_existent_template.xlsx")

    def test_invalid_day_raises_error(self):
        """Verify ValueError is raised if requested day does not exist in template."""
        report_data = self._build_dummy_report_data()
        with self.assertRaises(ValueError):
            generate_daily_excel(4, 2026, 32, report_data, template_path=self.template_path)


if __name__ == '__main__':
    unittest.main()
