import unittest
from core.calculator import Calculator
from core.gsheets_db import COMMODITIES, DIVISIONS, MONTH_DAYS

class MockGSheetsDB:
    """Mock database that returns controlled values for calculation testing."""
    def __init__(self, data_store=None):
        # data_store: dict mapping (tab, category, name, month) -> list of daily values
        self.data_store = data_store or {}

    def get_wagon_count(self, fy_tab, category, name, month, day):
        key = (fy_tab, category, name, month)
        if key in self.data_store and len(self.data_store[key]) >= day:
            return float(self.data_store[key][day - 1])
        return 0.0

    def get_month_sum(self, fy_tab, category, name, month, up_to_day=None):
        key = (fy_tab, category, name, month)
        if key not in self.data_store:
            return 0.0
        days_data = self.data_store[key]
        if up_to_day is not None:
            days_data = days_data[:up_to_day]
        return float(sum(days_data))

    def get_fy_progressive_sum(self, fy_tab, category, name, current_month, current_day):
        months_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
        total = 0.0
        for m in months_order:
            if m == current_month:
                total += self.get_month_sum(fy_tab, category, name, m, up_to_day=current_day)
                break
            else:
                total += self.get_month_sum(fy_tab, category, name, m)
        return total


class TestCalculator(unittest.TestCase):
    def setUp(self):
        self.db = MockGSheetsDB()
        self.calc = Calculator(self.db)

    def test_april_1st_fy_start_math(self):
        """On April 1st (Start of Financial Year):
        - Current MTD Avg MUST equal today's wagons (day 1)
        - Current FY Progressive Avg MUST equal today's wagons (day 1)
        - Prior months progressive loading MUST be strictly 0.0
        - Cumulative division loading MUST equal today's loading for each division
        """
        daily_input = {
            'CEMT': {'rakes': 9.0, 'wagons': 420.0},
            'COAL': {'rakes': 2.0, 'wagons': 110.0},
            'ADI': {'rakes': 4.0, 'wagons': 200.0},
            'GIMB': {'rakes': 5.0, 'wagons': 330.0},
            'BCT': {'rakes': 1.0, 'wagons': 50.0},
        }

        report = self.calc.calculate_daily_report(
            current_year_tab="FY 2026-27",
            prior_year_tab="FY 2025-26",
            month=4,
            day=1,
            daily_input=daily_input
        )

        # 1. Commodity level on Day 1
        cemt = report['commodity']['CEMT']
        self.assertEqual(cemt['wagons'], 420.0)
        self.assertEqual(cemt['cur_mtd_avg'], 420.0, "Day 1 MTD avg must equal today's wagons")
        self.assertEqual(cemt['cur_fy_avg'], 420.0, "Day 1 FY avg must equal today's wagons")

        coal = report['commodity']['COAL']
        self.assertEqual(coal['cur_mtd_avg'], 110.0)
        self.assertEqual(coal['cur_fy_avg'], 110.0)

        # 2. Right-hand summary panel on April 1st
        summary = report['summary']
        expected_today_wagons = 200.0 + 330.0 + 50.0  # ADI + GIMB + BCT
        self.assertEqual(summary['cur_today_wagons'], expected_today_wagons)
        self.assertEqual(summary['cur_mtd_wagons'], expected_today_wagons)
        self.assertEqual(summary['cur_fy_prog_wagons'], expected_today_wagons)

        # Crucial check: Zero prior months in new FY
        self.assertEqual(summary['cur_prev_months_wagons'], 0.0, "April must have 0 prior months loading in FY")
        self.assertEqual(summary['cur_prev_months_mt'], 0.0, "April must have 0.00 MT prior loading")
        self.assertEqual(summary['ly_prev_months_wagons'], 0.0)

        # Division cumulative check on April 1st
        div_cumms = summary['div_cumms']
        self.assertEqual(div_cumms['ADI_GIMB']['cur'], 530.0, "ADI + GIMB combined on April 1 must be 200 + 330 = 530")
        self.assertEqual(div_cumms['BCT']['cur'], 50.0)
        self.assertEqual(div_cumms['BRC']['cur'], 0.0)

    def test_fiscal_day_count_across_all_months(self):
        """Verify fiscal days since April 1st for key milestone dates."""
        milestones = [
            (4, 1, 1),      # April 1 -> Day 1
            (4, 30, 30),    # April 30 -> Day 30
            (5, 1, 31),     # May 1 -> Day 31
            (5, 31, 61),    # May 31 -> Day 61
            (6, 30, 91),    # June 30 -> Day 91
            (7, 31, 122),   # July 31 -> Day 122
            (8, 1, 123),    # August 1 -> Day 123
            (8, 15, 137),   # August 15 -> Day 137
            (9, 30, 183),   # September 30 -> Day 183
            (10, 31, 214),  # October 31 -> Day 214
            (11, 30, 244),  # November 30 -> Day 244
            (12, 31, 275),  # December 31 -> Day 275
            (1, 1, 276),    # January 1 -> Day 276
            (1, 31, 306),   # January 31 -> Day 306
            (2, 28, 334),   # February 28 -> Day 334
            (3, 31, 365),   # March 31 -> Day 365
        ]
        for month, day, expected_days in milestones:
            with self.subTest(month=month, day=day):
                actual = self.calc._get_fiscal_day_count(month, day)
                self.assertEqual(actual, expected_days, f"Fiscal day count for {day}/{month} should be {expected_days}, got {actual}")

    def test_mid_year_progressive_accumulation(self):
        """For August 2nd (Day 2 of Month 5 in FY):
        - Prior months must sum April + May + June + July (122 days)
        - Progressive wagons must equal prior months total + August MTD sum
        """
        # Populate mock db with 100 wagons/day for ADI across April, May, June, July
        mock_data = {
            ('FY 2026-27', 'division', 'ADI', 4): [100.0] * 30,  # 3000
            ('FY 2026-27', 'division', 'ADI', 5): [100.0] * 31,  # 3100
            ('FY 2026-27', 'division', 'ADI', 6): [100.0] * 30,  # 3000
            ('FY 2026-27', 'division', 'ADI', 7): [100.0] * 31,  # 3100
            ('FY 2026-27', 'division', 'ADI', 8): [120.0],        # Day 1 in Aug = 120
        }
        db = MockGSheetsDB(mock_data)
        calc = Calculator(db)

        # Day 2 input: 150 wagons today
        daily_input = {'ADI': {'rakes': 3.0, 'wagons': 150.0}}

        report = calc.calculate_daily_report(
            current_year_tab="FY 2026-27",
            prior_year_tab="FY 2025-26",
            month=8,
            day=2,
            daily_input=daily_input
        )

        adi = report['division']['ADI']
        # MTD sum for Aug 1-2 = 120 (yesterday) + 150 (today) = 270. Avg = 270 / 2 = 135
        self.assertEqual(adi['cur_mtd_avg'], 135.0)

        # Prior months total for ADI = 3000 + 3100 + 3000 + 3100 = 12200
        # Total FY sum = 12200 + 270 = 12470
        # Fiscal days on Aug 2 = 124 days
        # Cur FY avg = 12470 / 124 = 100.56
        self.assertEqual(adi['cur_fy_avg'], 100.56)

        # Check summary panel prior months
        summary = report['summary']
        self.assertEqual(summary['cur_prev_months_wagons'], 12200.0)
        self.assertEqual(summary['cur_mtd_wagons'], 270.0)
        self.assertEqual(summary['cur_fy_prog_wagons'], 12470.0)

    def test_diff_calculations_accuracy(self):
        """Verify diff in MTD average and diff in whole month average."""
        mock_data = {
            ('FY 2025-26', 'commodity', 'CEMT', 4): [400.0] * 30,  # Last year April had 400 wagons every day
        }
        db = MockGSheetsDB(mock_data)
        calc = Calculator(db)

        # This year April 1: loaded 450 wagons
        daily_input = {'CEMT': {'rakes': 9.0, 'wagons': 450.0}}
        report = calc.calculate_daily_report('FY 2026-27', 'FY 2025-26', 4, 1, daily_input)

        cemt = report['commodity']['CEMT']
        self.assertEqual(cemt['cur_mtd_avg'], 450.0)
        self.assertEqual(cemt['ly_wagons'], 400.0)
        self.assertEqual(cemt['ly_mtd_avg'], 400.0)
        # Diff MTD = 450 - 400 = +50
        self.assertEqual(cemt['diff_avg'], 50.0)
        # LY whole month avg = (400 * 30) / 30 = 400.0
        self.assertEqual(cemt['ly_whole_month_avg'], 400.0)
        # Diff whole month avg = 450 - 400 = +50
        self.assertEqual(cemt['diff_whole_month_avg'], 50.0)

    def test_negative_diffs(self):
        """Verify diff calculations when performance drops compared to last year."""
        mock_data = {
            ('FY 2025-26', 'commodity', 'COAL', 5): [300.0] * 31,
        }
        db = MockGSheetsDB(mock_data)
        calc = Calculator(db)

        # Loading only 200 wagons today
        daily_input = {'COAL': {'rakes': 4.0, 'wagons': 200.0}}
        report = calc.calculate_daily_report('FY 2026-27', 'FY 2025-26', 5, 1, daily_input)

        coal = report['commodity']['COAL']
        # Diff MTD = 200 - 300 = -100
        self.assertEqual(coal['diff_avg'], -100.0)
        self.assertEqual(coal['diff_whole_month_avg'], -100.0)

    def test_all_commodities_and_divisions_present(self):
        """Verify every commodity and division is returned in report_data."""
        daily_input = {}
        report = self.calc.calculate_daily_report('FY 2026-27', 'FY 2025-26', 4, 1, daily_input)

        for c in COMMODITIES:
            self.assertIn(c, report['commodity'])
            self.assertEqual(report['commodity'][c]['wagons'], 0.0)
            self.assertEqual(report['commodity'][c]['cur_mtd_avg'], 0.0)

        for d in DIVISIONS:
            self.assertIn(d, report['division'])
            self.assertEqual(report['division'][d]['wagons'], 0.0)
            self.assertEqual(report['division'][d]['cur_mtd_avg'], 0.0)

        self.assertIn('summary', report)

    def test_march_31st_fy_end_math(self):
        """On March 31st (Day 365 of FY):
        - Fiscal day count must be 365
        - All 11 prior months (April through February) must be summed in cur_prev_months_wagons
        """
        # Create prior months data
        mock_data = {}
        months_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2]
        for m in months_order:
            days = MONTH_DAYS[m]
            mock_data[('FY 2026-27', 'division', 'ADI', m)] = [10.0] * days

        # March 1 to 30 = 10 wagons/day
        mock_data[('FY 2026-27', 'division', 'ADI', 3)] = [10.0] * 30

        db = MockGSheetsDB(mock_data)
        calc = Calculator(db)

        daily_input = {'ADI': {'rakes': 1.0, 'wagons': 10.0}}
        report = calc.calculate_daily_report('FY 2026-27', 'FY 2025-26', 3, 31, daily_input)

        adi = report['division']['ADI']
        # Total fiscal days = 365
        # Total wagons = 365 * 10 = 3650
        # Cur FY avg = 3650 / 365 = 10.0
        self.assertEqual(adi['cur_fy_avg'], 10.0)
        self.assertEqual(adi['cur_mtd_avg'], 10.0)

        summary = report['summary']
        # Prior 11 months days = 334 days * 10 wagons = 3340
        self.assertEqual(summary['cur_prev_months_wagons'], 3340.0)
        # MTD March wagons = 31 * 10 = 310
        self.assertEqual(summary['cur_mtd_wagons'], 310.0)
        # Full FY progressive wagons = 3340 + 310 = 3650
        self.assertEqual(summary['cur_fy_prog_wagons'], 3650.0)

    def test_million_tonnes_conversion(self):
        """Verify Million Tonnes calculation: (wagons * 55.0) / 1,000,000 rounded to 2 decimals."""
        # 10,000 wagons * 55 tonnes = 550,000 tonnes = 0.55 MT
        daily_input = {
            'ADI': {'rakes': 100.0, 'wagons': 10000.0}
        }
        report = self.calc.calculate_daily_report('FY 2026-27', 'FY 2025-26', 4, 1, daily_input)
        summary = report['summary']
        self.assertEqual(summary['cur_today_wagons'], 10000.0)
        self.assertEqual(summary['cur_today_mt'], 0.55)

    def test_division_cumms_mapping_adi_gimb_and_bvc(self):
        """Verify ADI_GIMB is combined ADI + GIMB, and BVC receives BVP data."""
        daily_input = {
            'ADI': {'rakes': 2.0, 'wagons': 100.0},
            'GIMB': {'rakes': 3.0, 'wagons': 150.0},
            'BVP': {'rakes': 1.0, 'wagons': 60.0},
        }
        report = self.calc.calculate_daily_report('FY 2026-27', 'FY 2025-26', 4, 1, daily_input)
        div_cumms = report['summary']['div_cumms']

        self.assertEqual(div_cumms['ADI_GIMB']['cur'], 250.0, "ADI (100) + GIMB (150) must combine to 250")
        self.assertEqual(div_cumms['BVC']['cur'], 60.0, "BVC must receive Bhavnagar BVP (60) data")


if __name__ == '__main__':
    unittest.main()
