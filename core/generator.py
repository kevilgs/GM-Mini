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
        ws.cell(row=row, column=13).value = data['ly_full_year_avg'] # M
        
    # Calculate Totals
    # Openpyxl doesn't evaluate formulas immediately, so we can write sum formulas
    # or calculate the sums in Python. We will write formulas for the totals.
    ws.cell(row=18, column=3).value = "=SUM(C5:C17)"
    ws.cell(row=18, column=4).value = "=SUM(D5:D17)"
    ws.cell(row=18, column=5).value = "=SUM(E5:E17)"
    ws.cell(row=18, column=6).value = "=SUM(F5:F17)"
    ws.cell(row=18, column=7).value = "=SUM(G5:G17)"
    ws.cell(row=18, column=8).value = "=SUM(H5:H17)"
    ws.cell(row=18, column=9).value = "=SUM(I5:I17)"
    ws.cell(row=18, column=10).value = "=SUM(J5:J17)"
    ws.cell(row=18, column=11).value = "=SUM(K5:K17)"
    ws.cell(row=18, column=12).value = "=SUM(L5:L17)"
    ws.cell(row=18, column=13).value = "=SUM(M5:M17)"
    
    ws.cell(row=28, column=3).value = "=SUM(C21:C27)"
    ws.cell(row=28, column=4).value = "=SUM(D21:D27)"
    ws.cell(row=28, column=5).value = "=SUM(E21:E27)"
    ws.cell(row=28, column=6).value = "=SUM(F21:F27)"
    ws.cell(row=28, column=7).value = "=SUM(G21:G27)"
    ws.cell(row=28, column=8).value = "=SUM(H21:H27)"
    ws.cell(row=28, column=9).value = "=SUM(I21:I27)"
    ws.cell(row=28, column=10).value = "=SUM(J21:J27)"
    ws.cell(row=28, column=11).value = "=SUM(K21:K27)"
    ws.cell(row=28, column=12).value = "=SUM(L21:L27)"
    ws.cell(row=28, column=13).value = "=SUM(M21:M27)"
        
    wb.save(output_path)
    wb.close()
    
    return output_path
