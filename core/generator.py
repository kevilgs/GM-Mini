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
    Saves and returns the output path.
    """
    os.makedirs("generated_workbooks", exist_ok=True)
    month_name = datetime.date(year, month, 1).strftime("%B").upper()
    output_filename = f"{month_name}_{year}.xlsx"
    output_path = os.path.join("generated_workbooks", output_filename)
    
    file_to_open = output_path if os.path.exists(output_path) else template_path
    
    if not os.path.exists(file_to_open):
        raise FileNotFoundError(f"Source file {file_to_open} not found.")
        
    wb = openpyxl.load_workbook(file_to_open)
    
    day_str = str(day)
    if day_str not in wb.sheetnames:
        raise ValueError(f"Sheet for day '{day}' does not exist in template.")
        
    ws = wb[day_str]
    
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
