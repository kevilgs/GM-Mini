import openpyxl
import os
import datetime

# Row mappings in Excel Daily Sheet
EXCEL_COMMODITY_ROWS = {
    'CEMT': 5, 'COAL': 6, 'CONT': 7, 'FERT': 8, 'IMFT': 9, 'POL': 10, 
    'SALT': 11, 'STEEL': 12, 'DOC': 13, 'FG': 14, 'CHEM': 15, 'AUTO': 16, 'OTHERS': 17
}

EXCEL_DIVISION_ROWS = {
    'ADI': 21, 'GIMB': 22, 'BCT': 23, 'BRC': 24, 'RJT': 25, 'BVP': 26, 'RTM': 27
}

def generate_daily_excel(month: int, year: int, day: int, report_data: dict, template_path: str = os.path.join("assets", "GM MINI- AUGUST 2026.xlsx")):
    """
    Injects the calculator results into the Excel template for the specific day.
    Deletes all other garbage/empty sheets and saves a clean daily report.
    Returns the output path.
    """
    os.makedirs("generated_workbooks", exist_ok=True)
    output_filename = f"{day:02d}-{month:02d}-{year}.xlsx"
    output_path = os.path.join("generated_workbooks", output_filename)
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Source template file {template_path} not found.")
        
    wb = openpyxl.load_workbook(template_path)
    
    day_str = str(day)
    if day_str not in wb.sheetnames:
        raise ValueError(f"Sheet for day '{day}' does not exist in template.")
        
    # Delete all other daily sheets except the one for the current day.
    # We MUST preserve sheets like 'LINK' and 'Sheet1' because the formulas rely on them.
    for i in range(1, 32):
        s_name = str(i)
        if s_name != day_str and s_name in wb.sheetnames:
            del wb[s_name]
            
    ws = wb[day_str]
    # Optionally rename the single remaining tab to the exact date for maximum clarity
    ws.title = f"{day:02d}-{month:02d}-{year}"
    
    # Update dynamic headers
    target_date = datetime.date(year, month, day)
    ws.cell(row=1, column=16).value = target_date  # DATE
    
    # Calculate financial year and previous year
    fy_start_year = year if month >= 4 else year - 1
    fy_str = f"{fy_start_year}-{str(fy_start_year + 1)[-2:]}"
    prev_fy_str = f"{fy_start_year - 1}-{str(fy_start_year)[-2:]}"
    
    month_name = target_date.strftime("%b").upper()
    full_month_name = target_date.strftime("%B").upper()
    
    # If month is April, "Up to" is MARCH of previous year. Otherwise it's the previous month.
    prev_month_date = (target_date.replace(day=1) - datetime.timedelta(days=1))
    prev_month_name = prev_month_date.strftime("%B").upper()
    
    ws.cell(row=2, column=14).value = f"LOADING TARGET   {fy_str}"
    ws.cell(row=3, column=2).value = f"Target per day for  {month_name}-{str(year)[-2:]}"
    ws.cell(row=3, column=9).value = f"Avg  for the whole Month\n{month_name}-{year - 1}"
    ws.cell(row=3, column=11).value = f"Current Year   from April-{fy_start_year}"
    ws.cell(row=3, column=12).value = f"Last Year   from  April-{fy_start_year - 1}"
    ws.cell(row=3, column=13).value = f"Last Year Upto                    31st March-{fy_start_year}"
    ws.cell(row=3, column=14).value = target_date.replace(day=1) # 1st of month
    
    ws.cell(row=5, column=14).value = f"Loading Target {fy_start_year}-{fy_start_year + 1}"
    ws.cell(row=6, column=14).value = f"Progressive Loading Up To {prev_month_name}"
    
    ws.cell(row=19, column=9).value = f"Avg  for the whole Month\n{month_name}-{year - 1}"
    ws.cell(row=19, column=11).value = f"Current Year   from April-{fy_start_year}"
    ws.cell(row=19, column=12).value = f"Last Year   from  April-{fy_start_year - 1}"
    ws.cell(row=19, column=13).value = f"Last Year Upto                    31st March-{fy_start_year}"
    
    # Process Commodities
    for name, data in report_data['commodity'].items():
        row = EXCEL_COMMODITY_ROWS[name]
        ws.cell(row=row, column=3).value = data['rakes']          # C
        ws.cell(row=row, column=4).value = data['wagons']         # D
        ws.cell(row=row, column=5).value = data['cur_mtd_avg']    # E
        ws.cell(row=row, column=6).value = data['ly_wagons']      # F
        ws.cell(row=row, column=7).value = data['ly_mtd_avg']     # G
        ws.cell(row=row, column=8).value = data['diff_avg']       # H
        ws.cell(row=row, column=9).value = data['ly_whole_month_avg'] # I
        ws.cell(row=row, column=10).value = data['diff_whole_month_avg'] # J
        ws.cell(row=row, column=11).value = data['cur_fy_avg']    # K
        ws.cell(row=row, column=12).value = data['ly_fy_avg']     # L
        if data['ly_full_year_avg'] > 0:
            ws.cell(row=row, column=13).value = data['ly_full_year_avg'] # M

    # Process Divisions
    for name, data in report_data['division'].items():
        row = EXCEL_DIVISION_ROWS[name]
        ws.cell(row=row, column=3).value = data['rakes']          # C
        ws.cell(row=row, column=4).value = data['wagons']         # D
        ws.cell(row=row, column=5).value = data['cur_mtd_avg']    # E
        ws.cell(row=row, column=6).value = data['ly_wagons']      # F
        ws.cell(row=row, column=7).value = data['ly_mtd_avg']     # G
        ws.cell(row=row, column=8).value = data['diff_avg']       # H
        ws.cell(row=row, column=9).value = data['ly_whole_month_avg'] # I
        ws.cell(row=row, column=10).value = data['diff_whole_month_avg'] # J
        ws.cell(row=row, column=11).value = data['cur_fy_avg']    # K
        ws.cell(row=row, column=12).value = data['ly_fy_avg']     # L
        if data['ly_full_year_avg'] > 0:
            ws.cell(row=row, column=13).value = data['ly_full_year_avg'] # M
        
    # Calculate Totals for Tables
    for col in range(3, 14):
        col_letter = openpyxl.utils.get_column_letter(col)
        ws.cell(row=18, column=col).value = f"=SUM({col_letter}5:{col_letter}17)"
        ws.cell(row=28, column=col).value = f"=SUM({col_letter}21:{col_letter}27)"

    # --- Inject Dynamic Right-Hand Panel Metrics ---
    summary = report_data.get('summary', {})
    if summary:
        # Row 4 Year Labels
        ws.cell(row=4, column=16).value = prev_fy_str  # P4
        ws.cell(row=4, column=17).value = fy_str       # Q4
        ws.cell(row=4, column=18).value = prev_fy_str  # R4
        ws.cell(row=4, column=19).value = fy_str       # S4

        # Row 6: Progressive Loading Up To Last Month
        ws.cell(row=6, column=16).value = summary.get('ly_prev_months_wagons', 0.0)  # P6
        ws.cell(row=6, column=17).value = summary.get('cur_prev_months_wagons', 0.0) # Q6
        ws.cell(row=6, column=18).value = summary.get('ly_prev_months_mt', 0.0)      # R6
        ws.cell(row=6, column=19).value = summary.get('cur_prev_months_mt', 0.0)      # S6

        # Row 7: Today's Loading
        ws.cell(row=7, column=16).value = summary.get('ly_today_wagons', 0.0)         # P7
        ws.cell(row=7, column=17).value = summary.get('cur_today_wagons', 0.0)        # Q7
        ws.cell(row=7, column=18).value = summary.get('ly_today_mt', 0.0)             # R7
        ws.cell(row=7, column=19).value = summary.get('cur_today_mt', 0.0)             # S7

        # Row 8: Current Month Loading Till Date
        ws.cell(row=8, column=16).value = summary.get('ly_mtd_wagons', 0.0)           # P8
        ws.cell(row=8, column=17).value = summary.get('cur_mtd_wagons', 0.0)          # Q8
        ws.cell(row=8, column=18).value = summary.get('ly_mtd_mt', 0.0)               # R8
        ws.cell(row=8, column=19).value = summary.get('cur_mtd_mt', 0.0)               # S8

        # Rows 10-15: Division-Wise Cumulative Loading
        div_cumms = summary.get('div_cumms', {})
        div_rows = {
            'ADI_GIMB': 10,
            'BCT': 11,
            'BRC': 12,
            'RJT': 13,
            'BVC': 14,
            'RTM': 15
        }
        for div_key, r in div_rows.items():
            c_data = div_cumms.get(div_key, {'cur': 0.0, 'ly': 0.0})
            ws.cell(row=r, column=16).value = c_data['ly']  # P
            ws.cell(row=r, column=17).value = c_data['cur'] # Q
            # Convert cumulative wagons to MT (~55 tonnes/wagon)
            ly_mt = round((c_data['ly'] * 55.0) / 1_000_000, 2)
            cur_mt = round((c_data['cur'] * 55.0) / 1_000_000, 2)
            ws.cell(row=r, column=18).value = ly_mt         # R
            ws.cell(row=r, column=19).value = cur_mt        # S

        # Row 16: Progressive LDG from April
        ws.cell(row=16, column=16).value = summary.get('ly_fy_prog_wagons', 0.0)      # P16
        ws.cell(row=16, column=17).value = summary.get('cur_fy_prog_wagons', 0.0)     # Q16
        ws.cell(row=16, column=18).value = "=SUM(R10:R15)"
        ws.cell(row=16, column=19).value = "=SUM(S10:S15)"

        # Row 17: To Achieve Target
        target_wagons = ws.cell(row=5, column=17).value
        try:
            target_wagons_num = float(target_wagons)
        except (ValueError, TypeError):
            target_wagons_num = 2208615.0
            
        cur_prog = summary.get('cur_fy_prog_wagons', 0.0)
        ws.cell(row=17, column=17).value = round(target_wagons_num - cur_prog, 2)    # Q17
        ws.cell(row=17, column=19).value = "=+S5-S16"                                # S17

        # Clear hardcoded August rake stock numbers (Rows 21-27)
        for sr in range(21, 28):
            ws.cell(row=sr, column=16).value = 0 # P
            ws.cell(row=sr, column=18).value = 0 # R
        ws.cell(row=28, column=16).value = "=SUM(P21:P27)"
        ws.cell(row=28, column=18).value = "=SUM(R21:R27)"

    # Enable full automatic calculation on workbook open
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
        
    wb.save(output_path)
    wb.close()
    
    return output_path
