import os
import datetime
from flask import Flask, render_template, request, jsonify, send_file, flash, redirect, url_for
from core.calculator import Calculator
from core.gsheets_db import GSheetsDB, COMMODITIES, DIVISIONS
from core.generator import generate_daily_excel
from core.drive_api import upload_excel_to_drive
from googleapiclient.errors import HttpError

app = Flask(__name__)
app.secret_key = 'super_secret_key'

db = GSheetsDB()
calc = Calculator(db)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html', 
                           commodities=COMMODITIES,
                           divisions=DIVISIONS)

@app.route('/api/check_date', methods=['POST'])
def check_date():
    date_str = request.form.get('report_date')
    if not date_str:
        return jsonify({"status": "error", "message": "No date provided."})
        
    try:
        report_date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({"status": "error", "message": "Invalid date format."})
        
    month = report_date.month
    year = report_date.year
    fy_start_year = year if month >= 4 else year - 1
    current_year_tab = f"FY {fy_start_year}-{str(fy_start_year + 1)[-2:]}"
    
    last_filled = db.get_last_filled_date(current_year_tab)
    
    if last_filled is None:
        # DB is empty for this FY, must start at April 1st
        expected_start = datetime.date(fy_start_year, 4, 1)
        if report_date != expected_start:
            return jsonify({
                "status": "error", 
                "message": f"Database is empty for {current_year_tab}. You must start by filling data for {expected_start.strftime('%d-%m-%Y')}."
            })
        return jsonify({"status": "ok"})
        
    # Check chronological constraints
    next_expected = last_filled + datetime.timedelta(days=1)
    
    if report_date > next_expected:
        return jsonify({
            "status": "error",
            "message": f"Cannot skip dates. The last filled date is {last_filled.strftime('%d-%m-%Y')}. Please fill data for {next_expected.strftime('%d-%m-%Y')} next."
        })
    elif report_date <= last_filled:
        return jsonify({
            "status": "warning",
            "message": f"Data for {report_date.strftime('%d-%m-%Y')} is already present in the database. Do you want to overwrite it with new data?"
        })
        
    return jsonify({"status": "ok"})

@app.route('/generate', methods=['POST'])
def generate():
    try:
        date_str = request.form.get('report_date')
        if not date_str:
            flash("Please select a date.")
            return redirect(url_for('index'))
            
        report_date = datetime.datetime.strptime(date_str, '%Y-%m-%d')
        month = report_date.month
        day = report_date.day
        year = report_date.year
        
        # Determine the FY tab string for current year and prior year.
        # Logic: If month is April-Dec, FY is Year to Year+1. If Jan-Mar, FY is Year-1 to Year.
        fy_start_year = year if month >= 4 else year - 1
        current_year_tab = f"FY {fy_start_year}-{str(fy_start_year + 1)[-2:]}"
        prior_year_tab = f"FY {fy_start_year - 1}-{str(fy_start_year)[-2:]}"
        
        # Build daily input dictionary from form data
        daily_input = {}
        for name in COMMODITIES + DIVISIONS:
            rakes_val = request.form.get(f"{name}_rakes", "0")
            wagons_val = request.form.get(f"{name}_wagons", "0")
            
            try:
                rakes = float(rakes_val) if rakes_val else 0.0
                wagons = float(wagons_val) if wagons_val else 0.0
            except ValueError:
                rakes = 0.0
                wagons = 0.0
                
            daily_input[name] = {'rakes': rakes, 'wagons': wagons}
            
        print(f"Generating report for {day}/{month}/{year}")
        print(f"Current FY Tab: {current_year_tab}, Prior FY Tab: {prior_year_tab}")
        # --- STRICT BACKEND DATE VALIDATION ---
        last_filled = db.get_last_filled_date(current_year_tab)
        
        if last_filled is None:
            expected_start = datetime.date(fy_start_year, 4, 1)
            if report_date.date() != expected_start:
                flash(f"Database is empty for {current_year_tab}. You must start by filling data for {expected_start.strftime('%d-%m-%Y')}.")
                return redirect(url_for('index'))
        else:
            next_expected = last_filled + datetime.timedelta(days=1)
            if report_date.date() > next_expected:
                flash(f"Cannot skip dates! The last filled date is {last_filled.strftime('%d-%m-%Y')}. Please fill data for {next_expected.strftime('%d-%m-%Y')} next.")
                return redirect(url_for('index'))
            elif report_date.date() <= last_filled:
                if request.form.get('force_overwrite') == 'true':
                    pass # User confirmed overwrite via JS popup
                else:
                    flash(f"Data for {report_date.strftime('%d-%m-%Y')} is already present in the database. Overwriting is blocked without confirmation.")
                    return redirect(url_for('index'))
        # --------------------------------------
        
        # 1. Update Database with today's entries
        db.save_daily_batch(current_year_tab, month, day, daily_input)
        
        # 2. Calculate Dashboard Math
        report_data = calc.calculate_daily_report(current_year_tab, prior_year_tab, month, day, daily_input)
        
        # 3. Generate Excel
        output_file = generate_daily_excel(month, year, day, report_data)
        
        # 4. Upload to Google Drive
        month_name = datetime.date(year, month, 1).strftime('%B-%Y')
        filename = f"{day:02d}-{month:02d}-{year}.xlsx"
        
        link = upload_excel_to_drive(output_file, filename, current_year_tab, month_name)
        
        flash(f"Success! Excel report for {day:02d}-{month:02d}-{year} has been generated and uploaded to Google Drive. <br><a href='{link}' target='_blank' style='color: var(--color-success); text-decoration: underline;'>Click here to view it.</a>", "success")
        return redirect(url_for('index'))
        
    except HttpError as e:
        import traceback
        traceback.print_exc()
        error_details = e.error_details[0] if getattr(e, 'error_details', None) and len(e.error_details) > 0 else {}
        if error_details.get('reason') == 'storageQuotaExceeded':
            flash("Error: The service account has run out of storage space. This usually happens if it created the 'GM-MINI SHEETS' folder itself. Please delete any 'GM-MINI SHEETS' folders, then create one in your personal Google Drive and share it with the service account as an Editor.", "error")
        else:
            flash(f"Google Drive API Error: {str(e)}", "error")
        return redirect(url_for('index'))
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Error: {str(e)}", "error")
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=5050)
