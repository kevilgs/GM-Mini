from core.gsheets_db import GSheetsDB, COMMODITIES, DIVISIONS, MONTH_DAYS
import datetime

class Calculator:
    def __init__(self, db: GSheetsDB):
        self.db = db
        
    def calculate_daily_report(self, current_year_tab, prior_year_tab, month, day, daily_input):
        """
        daily_input: dict mapping commodity/division to {'rakes': X, 'wagons': Y}
        Example: {'CEMT': {'rakes': 9, 'wagons': 469}, 'ADI': {'rakes': 5, 'wagons': 250}, ...}
        
        Returns a dict of all calculated values for the Excel template.
        """
        
        report_data = {
            'commodity': {},
            'division': {}
        }
        
        fiscal_day_count = self._get_fiscal_day_count(month, day)
        
        for category, items in [('commodity', COMMODITIES), ('division', DIVISIONS)]:
            for name in items:
                # 1. Inputs
                input_data = daily_input.get(name, {'rakes': 0.0, 'wagons': 0.0})
                cur_wagons = float(input_data['wagons'])
                cur_rakes = float(input_data['rakes'])
                
                # 2. Current Month-To-Date (MTD) Avg
                # Get sum from day 1 to yesterday from DB, then add today's input
                if day > 1:
                    sum_till_yesterday = self.db.get_month_sum(current_year_tab, category, name, month, up_to_day=day-1)
                else:
                    sum_till_yesterday = 0.0
                    
                cur_mtd_sum = sum_till_yesterday + cur_wagons
                cur_mtd_avg = cur_mtd_sum / day
                
                # 3. Last Year Same Day Wagons
                ly_wagons = self.db.get_wagon_count(prior_year_tab, category, name, month, day)
                
                # 4. Last Year MTD Avg
                ly_mtd_sum = self.db.get_month_sum(prior_year_tab, category, name, month, up_to_day=day)
                ly_mtd_avg = ly_mtd_sum / day
                
                # 5. Diff in AVG (MTD)
                diff_avg = cur_mtd_avg - ly_mtd_avg
                
                # 6. Last Year Whole Month Avg
                # (Sum of all days in that month last year / max days)
                ly_whole_month_sum = self.db.get_month_sum(prior_year_tab, category, name, month)
                ly_whole_month_avg = ly_whole_month_sum / MONTH_DAYS[month]
                
                # 7. Diff in Avg (Current MTD vs Last Year Whole Month)
                diff_whole_month_avg = cur_mtd_avg - ly_whole_month_avg
                
                # 8. Current FY Progressive Avg
                # Get total from April 1 to last month, plus current MTD sum
                fy_till_last_month = 0.0
                months_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
                for m in months_order:
                    if m == month:
                        break
                    fy_till_last_month += self.db.get_month_sum(current_year_tab, category, name, m)
                    
                cur_fy_sum = fy_till_last_month + cur_mtd_sum
                cur_fy_avg = cur_fy_sum / fiscal_day_count
                
                # 9. Last Year FY Progressive Avg (up to exact same date)
                ly_fy_sum = self.db.get_fy_progressive_sum(prior_year_tab, category, name, month, day)
                ly_fy_avg = ly_fy_sum / fiscal_day_count
                
                # 10. Last Year Upto 31st March (Full Year Avg)
                ly_full_year_sum = self.db.get_fy_progressive_sum(prior_year_tab, category, name, 3, 31)
                ly_full_year_avg = ly_full_year_sum / 365
                
                report_data[category][name] = {
                    'rakes': cur_rakes,
                    'wagons': cur_wagons,
                    'cur_mtd_avg': round(cur_mtd_avg, 2),
                    'ly_wagons': round(ly_wagons, 2),
                    'ly_mtd_avg': round(ly_mtd_avg, 2),
                    'diff_avg': round(diff_avg, 2),
                    'ly_whole_month_avg': round(ly_whole_month_avg, 2),
                    'diff_whole_month_avg': round(diff_whole_month_avg, 2),
                    'cur_fy_avg': round(cur_fy_avg, 2),
                    'ly_fy_avg': round(ly_fy_avg, 2),
                    'ly_full_year_avg': round(ly_full_year_avg, 2)
                }
                
        return report_data
        
    def _get_fiscal_day_count(self, month, day):
        # Calculate days since April 1st
        months_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
        total_days = 0
        for m in months_order:
            if m == month:
                total_days += day
                break
            else:
                total_days += MONTH_DAYS[m]
        return total_days

if __name__ == "__main__":
    db = GSheetsDB()
    calc = Calculator(db)
    
    # Dummy input for August 5
    dummy_input = {
        'CEMT': {'rakes': 9, 'wagons': 469},
        'COAL': {'rakes': 3, 'wagons': 150}
    }
    
    # We use FY 2025-26 as current year for this test because we don't have FY 2026-27 populated yet.
    print("Testing calculator for August 5th...")
    result = calc.calculate_daily_report("FY 2025-26", "FY 2025-26", 8, 5, dummy_input)
    print("CEMT calculation:")
    for k, v in result['commodity']['CEMT'].items():
        print(f"  {k}: {v}")
