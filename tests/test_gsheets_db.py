import unittest
from unittest.mock import MagicMock, patch
import datetime
from core.gsheets_db import (
    GSheetsDB,
    CATEGORY_COLS,
    COMMODITIES,
    DIVISIONS,
    MONTH_DAYS,
    get_month_days
)

class TestGSheetsDBLogic(unittest.TestCase):
    def setUp(self):
        # Create an uninitialized GSheetsDB instance to test internal logic without network
        self.db = object.__new__(GSheetsDB)
        self.db._cache = {}
        self.db.spreadsheet = MagicMock()

    def test_category_cols_layout(self):
        """Verify the 22-column row layout matches the specification:
        - 13 Commodities in cols 1-13
        - Separator at col 14
        - 7 Divisions in cols 15-21
        """
        self.assertEqual(len(COMMODITIES), 13)
        self.assertEqual(len(DIVISIONS), 7)
        self.assertEqual(len(CATEGORY_COLS), 20)

        # Verify commodity indices
        expected_commodities = [
            'CEMT', 'COAL', 'CONT', 'FERT', 'IMFT', 'POL',
            'SALT', 'STEEL', 'DOC', 'FG', 'CHEM', 'AUTO', 'OTHERS'
        ]
        for idx, comm in enumerate(expected_commodities, start=1):
            self.assertEqual(CATEGORY_COLS[comm], idx, f"{comm} should be at column index {idx}")

        # Verify division indices (col 14 is blank separator)
        expected_divisions = ['ADI', 'GIMB', 'BCT', 'BRC', 'RJT', 'BVP', 'RTM']
        for idx, div in enumerate(expected_divisions, start=15):
            self.assertEqual(CATEGORY_COLS[div], idx, f"{div} should be at column index {idx}")

    def test_month_days_constant(self):
        """Verify MONTH_DAYS has all 12 months with correct days."""
        self.assertEqual(len(MONTH_DAYS), 12)
        self.assertEqual(MONTH_DAYS[4], 30)  # April
        self.assertEqual(MONTH_DAYS[5], 31)  # May
        self.assertEqual(MONTH_DAYS[6], 30)  # June
        self.assertEqual(MONTH_DAYS[7], 31)  # July
        self.assertEqual(MONTH_DAYS[8], 31)  # August
        self.assertEqual(MONTH_DAYS[9], 30)  # September
        self.assertEqual(MONTH_DAYS[10], 31) # October
        self.assertEqual(MONTH_DAYS[11], 30) # November
        self.assertEqual(MONTH_DAYS[12], 31) # December
        self.assertEqual(MONTH_DAYS[1], 31)  # January
        self.assertEqual(MONTH_DAYS[2], 28)  # February
        self.assertEqual(MONTH_DAYS[3], 31)  # March

    def test_dynamic_get_month_days_leap_year(self):
        """Verify get_month_days dynamically returns 29 for February in leap years and 28 otherwise."""
        self.assertEqual(get_month_days(2028, 2), 29, "February 2028 is a leap year (29 days)")
        self.assertEqual(get_month_days(2024, 2), 29, "February 2024 is a leap year (29 days)")
        self.assertEqual(get_month_days(2026, 2), 28, "February 2026 is a normal year (28 days)")
        self.assertEqual(get_month_days(2027, 2), 28, "February 2027 is a normal year (28 days)")
        self.assertEqual(get_month_days(2026, 4), 30, "April has 30 days")
        self.assertEqual(get_month_days(2026, 5), 31, "May has 31 days")

    def test_get_row_for_date(self):
        """Verify finding the correct row index given DD-MM-YYYY."""
        dummy_data = [
            ['Date', 'CEMT', 'COAL'],  # Row 0: header
            ['01-04-2026', '420', '110'], # Row 1
            ['02-04-2026', '430', '115'], # Row 2
            ['03-04-2026', '400', '105'], # Row 3
        ]
        self.assertEqual(self.db._get_row_for_date(dummy_data, '01-04-2026'), 1)
        self.assertEqual(self.db._get_row_for_date(dummy_data, '03-04-2026'), 3)
        self.assertEqual(self.db._get_row_for_date(dummy_data, '04-04-2026'), -1)

    def test_get_wagon_count(self):
        """Verify reading a specific day's wagon count from cached sheet data."""
        # Row 1: 01-04-2026
        # CEMT is col 1, ADI is col 15
        row1 = [""] * 22
        row1[0] = "01-04-2026"
        row1[CATEGORY_COLS['CEMT']] = "450"
        row1[CATEGORY_COLS['ADI']] = "210"

        header = ["Date"] + [""] * 21
        self.db._cache['FY 2026-27'] = [header, row1]

        # Valid values
        self.assertEqual(self.db.get_wagon_count('FY 2026-27', 'commodity', 'CEMT', 4, 1), 450.0)
        self.assertEqual(self.db.get_wagon_count('FY 2026-27', 'division', 'ADI', 4, 1), 210.0)

        # Empty or non-existent
        self.assertEqual(self.db.get_wagon_count('FY 2026-27', 'commodity', 'COAL', 4, 1), 0.0)
        self.assertEqual(self.db.get_wagon_count('FY 2026-27', 'commodity', 'CEMT', 4, 2), 0.0)

    def test_get_month_sum(self):
        """Verify monthly sum with and without up_to_day filter."""
        header = ["Date"] + [""] * 21
        rows = [header]
        # 5 days of data for April 2026: 100 wagons/day of CEMT
        for d in range(1, 6):
            row = [""] * 22
            row[0] = f"{d:02d}-04-2026"
            row[CATEGORY_COLS['CEMT']] = "100"
            rows.append(row)

        self.db._cache['FY 2026-27'] = rows

        # Full month sum = 5 * 100 = 500
        full_sum = self.db.get_month_sum('FY 2026-27', 'commodity', 'CEMT', 4)
        self.assertEqual(full_sum, 500.0)

        # Up to day 3 = 3 * 100 = 300
        partial_sum = self.db.get_month_sum('FY 2026-27', 'commodity', 'CEMT', 4, up_to_day=3)
        self.assertEqual(partial_sum, 300.0)

    def test_get_fy_progressive_sum(self):
        """Verify progressive summation across fiscal year months:
        April (4) -> May (5) -> ... -> March (3)
        """
        header = ["Date"] + [""] * 21
        rows = [header]
        # April: 2 days * 50 = 100
        for d in [1, 2]:
            r = [""] * 22
            r[0] = f"{d:02d}-04-2026"
            r[CATEGORY_COLS['ADI']] = "50"
            rows.append(r)
        # May: 2 days * 100 = 200
        for d in [1, 2]:
            r = [""] * 22
            r[0] = f"{d:02d}-05-2026"
            r[CATEGORY_COLS['ADI']] = "100"
            rows.append(r)

        self.db._cache['FY 2026-27'] = rows

        # Up to May 1st: All April (100) + May 1st (100) = 200
        prog_may1 = self.db.get_fy_progressive_sum('FY 2026-27', 'division', 'ADI', 5, 1)
        self.assertEqual(prog_may1, 200.0)

        # Up to May 2nd: All April (100) + May 1-2 (200) = 300
        prog_may2 = self.db.get_fy_progressive_sum('FY 2026-27', 'division', 'ADI', 5, 2)
        self.assertEqual(prog_may2, 300.0)

    def test_get_last_filled_date(self):
        """Verify bottom-up scan for the latest filled date."""
        header = ["Date"] + [""] * 21
        rows = [
            header,
            ["01-04-2026"] + [""] * 21,
            ["02-04-2026"] + [""] * 21,
            ["03-04-2026"] + [""] * 21,
            ["", "", ""], # empty trailing row
        ]
        self.db._cache['FY 2026-27'] = rows

        last_date = self.db.get_last_filled_date('FY 2026-27')
        self.assertEqual(last_date, datetime.date(2026, 4, 3))

    def test_get_last_filled_date_empty_sheet(self):
        """Verify None returned when sheet has no data rows."""
        header = ["Date"] + [""] * 21
        self.db._cache['FY 2026-27'] = [header]
        self.assertIsNone(self.db.get_last_filled_date('FY 2026-27'))

    def test_save_daily_batch_row_formatting(self):
        """Verify that save_daily_batch formats a 22-column row correctly."""
        mock_worksheet = MagicMock()
        self.db.spreadsheet.worksheet.return_value = mock_worksheet
        self.db._cache['FY 2026-27'] = [["Date"] + [""] * 21]

        daily_input = {
            'CEMT': {'wagons': '420', 'rakes': '9'},
            'COAL': {'wagons': '110', 'rakes': '2'},
            'ADI': {'wagons': '200', 'rakes': '4'},
        }

        success = self.db.save_daily_batch('FY 2026-27', 4, 1, daily_input)
        self.assertTrue(success)

        # Verify update was called on worksheet with 22 elements
        call_args = mock_worksheet.update.call_args
        self.assertIsNotNone(call_args)
        written_row = call_args.kwargs['values'][0]
        self.assertEqual(len(written_row), 22)
        self.assertEqual(written_row[0], "01-04-2026")
        self.assertEqual(written_row[CATEGORY_COLS['CEMT']], "420")
        self.assertEqual(written_row[CATEGORY_COLS['COAL']], "110")
        self.assertEqual(written_row[CATEGORY_COLS['ADI']], "200")
        self.assertEqual(written_row[14], "", "Separator column 14 must be empty")


if __name__ == '__main__':
    unittest.main()
