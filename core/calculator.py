from core.gsheets_db import GSheetsDB, COMMODITIES, DIVISIONS, MONTH_DAYS, get_month_days
import datetime
import calendar

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
        
        # Determine starting calendar years of current and prior financial years
        fy_start_year = int(current_year_tab.split('-')[0].replace("FY ", ""))
        prior_start_year = int(prior_year_tab.split('-')[0].replace("FY ", ""))
        
        fiscal_day_count = self._get_fiscal_day_count(month, day, fy_start_year=fy_start_year)
        
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
                # (Sum of all days in that month last year / days in that specific month)
                ly_whole_month_sum = self.db.get_month_sum(prior_year_tab, category, name, month)
                prior_cal_year = prior_start_year if month >= 4 else prior_start_year + 1
                ly_days_in_month = get_month_days(prior_cal_year, month)
                ly_whole_month_avg = ly_whole_month_sum / ly_days_in_month
                
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
                # In Indian FY, February falls in prior_start_year + 1
                ly_total_days = 366 if calendar.isleap(prior_start_year + 1) else 365
                ly_full_year_avg = ly_full_year_sum / ly_total_days
                
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

        # --- Calculate Right-Hand Summary Panel Metrics ---
        # 1. Today Totals
        cur_today_wagons = sum(float(daily_input.get(d, {}).get('wagons', 0.0)) for d in DIVISIONS)
        cur_today_rakes = sum(float(daily_input.get(d, {}).get('rakes', 0.0)) for d in DIVISIONS)
        ly_today_wagons = sum(report_data['division'][d]['ly_wagons'] for d in DIVISIONS)

        # 2. Month-To-Date Totals
        cur_mtd_wagons = sum(report_data['division'][d]['cur_mtd_avg'] * day for d in DIVISIONS)
        ly_mtd_wagons = sum(report_data['division'][d]['ly_mtd_avg'] * day for d in DIVISIONS)

        # 3. Progressive Loading Up To Last Month (within the financial year)
        # For April (month 4), there are 0 prior months in the new FY
        cur_prev_months_wagons = 0.0
        ly_prev_months_wagons = 0.0
        months_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
        for m in months_order:
            if m == month:
                break
            cur_prev_months_wagons += sum(self.db.get_month_sum(current_year_tab, 'division', d, m) for d in DIVISIONS)
            ly_prev_months_wagons += sum(self.db.get_month_sum(prior_year_tab, 'division', d, m) for d in DIVISIONS)

        # 4. Division-Wise Cumulative Loading from April 1 to Today
        div_cumms = {}
        for d in DIVISIONS:
            cur_d_cum = 0.0
            for m in months_order:
                if m == month:
                    if day > 1:
                        cur_d_cum += self.db.get_month_sum(current_year_tab, 'division', d, m, up_to_day=day-1)
                    cur_d_cum += float(daily_input.get(d, {}).get('wagons', 0.0))
                    break
                else:
                    cur_d_cum += self.db.get_month_sum(current_year_tab, 'division', d, m)

            ly_d_cum = self.db.get_fy_progressive_sum(prior_year_tab, 'division', d, month, day)
            div_cumms[d] = {'cur': cur_d_cum, 'ly': ly_d_cum}

        # ADI in GM MINI right panel combines ADI + GIMB
        adi_gimb_cur = div_cumms['ADI']['cur'] + div_cumms['GIMB']['cur']
        adi_gimb_ly = div_cumms['ADI']['ly'] + div_cumms['GIMB']['ly']

        cur_fy_prog_wagons = cur_prev_months_wagons + cur_mtd_wagons
        ly_fy_prog_wagons = ly_prev_months_wagons + ly_mtd_wagons

        # Million Tonnes: ~50-55 tonnes per wagon standard railway conversion
        tonnes_per_wagon = 55.0
        cur_today_mt = round((cur_today_wagons * tonnes_per_wagon) / 1_000_000, 2)
        ly_today_mt = round((ly_today_wagons * tonnes_per_wagon) / 1_000_000, 2)
        cur_mtd_mt = round((cur_mtd_wagons * tonnes_per_wagon) / 1_000_000, 2)
        ly_mtd_mt = round((ly_mtd_wagons * tonnes_per_wagon) / 1_000_000, 2)
        cur_prev_months_mt = round((cur_prev_months_wagons * tonnes_per_wagon) / 1_000_000, 2)
        ly_prev_months_mt = round((ly_prev_months_wagons * tonnes_per_wagon) / 1_000_000, 2)

        report_data['summary'] = {
            'cur_today_wagons': cur_today_wagons,
            'cur_today_rakes': cur_today_rakes,
            'ly_today_wagons': ly_today_wagons,
            'cur_today_mt': cur_today_mt,
            'ly_today_mt': ly_today_mt,
            'cur_mtd_wagons': cur_mtd_wagons,
            'ly_mtd_wagons': ly_mtd_wagons,
            'cur_mtd_mt': cur_mtd_mt,
            'ly_mtd_mt': ly_mtd_mt,
            'cur_prev_months_wagons': cur_prev_months_wagons,
            'ly_prev_months_wagons': ly_prev_months_wagons,
            'cur_prev_months_mt': cur_prev_months_mt,
            'ly_prev_months_mt': ly_prev_months_mt,
            'cur_fy_prog_wagons': cur_fy_prog_wagons,
            'ly_fy_prog_wagons': ly_fy_prog_wagons,
            'div_cumms': {
                'ADI_GIMB': {'cur': adi_gimb_cur, 'ly': adi_gimb_ly},
                'BCT': div_cumms['BCT'],
                'BRC': div_cumms['BRC'],
                'RJT': div_cumms['RJT'],
                'BVC': div_cumms['BVP'],
                'RTM': div_cumms['RTM'],
            }
        }
        
        return report_data
        
    def _get_fiscal_day_count(self, month, day, fy_start_year=None):
        # Calculate days since April 1st, accounting for leap years when fy_start_year is known
        months_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
        total_days = 0
        for m in months_order:
            if m == month:
                total_days += day
                break
            else:
                if fy_start_year:
                    m_year = fy_start_year if m >= 4 else fy_start_year + 1
                    total_days += get_month_days(m_year, m)
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
