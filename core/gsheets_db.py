import gspread
from google.oauth2.service_account import Credentials
import datetime

# Define column indices for the new 1-row-per-day layout (0-indexed in python list)
# Date is 0
CATEGORY_COLS = {
    'CEMT': 1, 'COAL': 2, 'CONT': 3, 'FERT': 4, 'IMFT': 5, 'POL': 6, 
    'SALT': 7, 'STEEL': 8, 'DOC': 9, 'FG': 10, 'CHEM': 11, 'AUTO': 12, 'OTHERS': 13,
    # 14 is empty separator
    'ADI': 15, 'GIMB': 16, 'BCT': 17, 'BRC': 18, 'RJT': 19, 'BVP': 20, 'RTM': 21
}

COMMODITIES = ['CEMT', 'COAL', 'CONT', 'FERT', 'IMFT', 'POL', 'SALT', 'STEEL', 'DOC', 'FG', 'CHEM', 'AUTO', 'OTHERS']
DIVISIONS = ['ADI', 'GIMB', 'BCT', 'BRC', 'RJT', 'BVP', 'RTM']

# The number of days in each month
MONTH_DAYS = {
    4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31, 1: 31, 2: 28, 3: 31
}

class GSheetsDB:
    def __init__(self, credentials_path='credentials.json', spreadsheet_name='GM-Mini-Database'):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
        self.client = gspread.authorize(creds)
        self.spreadsheet = self.client.open(spreadsheet_name)
        self._cache = {}
        
    def _get_worksheet_data(self, fy_tab, bypass_cache=False):
        """Fetches the entire worksheet as a 2D list with caching, automatically creating it if missing."""
        if fy_tab not in self._cache or bypass_cache:
            try:
                worksheet = self.spreadsheet.worksheet(fy_tab)
            except gspread.exceptions.WorksheetNotFound:
                print(f"Tab {fy_tab} not found. Creating automatically...")
                worksheet = self._create_and_format_worksheet(fy_tab)
            
            self._cache[fy_tab] = worksheet.get_all_values()
        return self._cache[fy_tab]
        
    def _create_and_format_worksheet(self, fy_tab):
        """Internal helper to create a new FY tab and apply the strict 22-column layout."""
        ws = self.spreadsheet.add_worksheet(title=fy_tab, rows="500", cols="22")
        
        header = [
            'Date', 
            'CEMT', 'COAL', 'CONT', 'FERT', 'IMFT', 'POL', 'SALT', 'STEEL', 'DOC', 'FG', 'CHEM', 'AUTO', 'OTHERS',
            '', 
            'ADI', 'GIMB', 'BCT', 'BRC', 'RJT', 'BVP', 'RTM'
        ]
        ws.update('A1:V1', [header])
        
        requests = []
        sheet_id = ws.id
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': sheet_id, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 22},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True},
                        'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE',
                        'borders': {
                            'bottom': {'style': 'SOLID'}, 'top': {'style': 'SOLID'}, 'left': {'style': 'SOLID'}, 'right': {'style': 'SOLID'}
                        }
                    }
                },
                'fields': 'userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment,borders)'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': sheet_id, 'startRowIndex': 1, 'endRowIndex': 500, 'startColumnIndex': 0, 'endColumnIndex': 1},
                'cell': {'userEnteredFormat': {'textFormat': {'bold': True}, 'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE'}},
                'fields': 'userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)'
            }
        })

        requests.append({
            'repeatCell': {
                'range': {'sheetId': sheet_id, 'startRowIndex': 1, 'endRowIndex': 500, 'startColumnIndex': 1, 'endColumnIndex': 22},
                'cell': {'userEnteredFormat': {'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE'}},
                'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment)'
            }
        })

        requests.append({
            'updateSheetProperties': {
                'properties': {'sheetId': sheet_id, 'gridProperties': {'frozenRowCount': 1, 'frozenColumnCount': 1}},
                'fields': 'gridProperties(frozenRowCount,frozenColumnCount)'
            }
        })

        requests.append({
            'updateDimensionProperties': {
                'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS', 'startIndex': 14, 'endIndex': 15},
                'properties': {'pixelSize': 30},
                'fields': 'pixelSize'
            }
        })

        requests.append({
            'deleteDimension': {
                'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS', 'startIndex': 22}
            }
        })

        self.spreadsheet.batch_update({'requests': requests})
        return ws
        
    def _get_row_for_date(self, data, target_date_str):
        # target_date_str is DD-MM-YYYY
        for row_idx in range(1, len(data)):
            if data[row_idx] and data[row_idx][0] == target_date_str:
                return row_idx
        return -1
        
    def get_wagon_count(self, fy_tab, category, name, month, day):
        data = self._get_worksheet_data(fy_tab)
        
        # Convert month, day to DD-MM-YYYY
        start_year = int(fy_tab.split('-')[0].replace("FY ", ""))
        year = start_year if month >= 4 else start_year + 1
        date_str = f"{day:02d}-{month:02d}-{year}"
        
        row_idx = self._get_row_for_date(data, date_str)
        if row_idx == -1:
            return 0.0
            
        col_idx = CATEGORY_COLS[name]
        if col_idx < len(data[row_idx]):
            val = data[row_idx][col_idx].strip()
            try:
                return float(val) if val != "" else 0.0
            except ValueError:
                return 0.0
        return 0.0
        
    def get_month_sum(self, fy_tab, category, name, month, up_to_day=None):
        data = self._get_worksheet_data(fy_tab)
        start_year = int(fy_tab.split('-')[0].replace("FY ", ""))
        year = start_year if month >= 4 else start_year + 1
        
        col_idx = CATEGORY_COLS[name]
        total = 0.0
        
        for row in data[1:]:
            if not row or not row[0]:
                continue
            date_str = row[0]
            try:
                d = datetime.datetime.strptime(date_str, '%d-%m-%Y')
                if d.year == year and d.month == month:
                    if up_to_day is None or d.day <= up_to_day:
                        if col_idx < len(row):
                            val = row[col_idx].strip()
                            if val:
                                total += float(val)
            except ValueError:
                continue
                
        return total
        
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

    def get_last_filled_date(self, fy_tab):
        data = self._get_worksheet_data(fy_tab)
        
        # Start from the bottom up to find the last valid date
        for row_idx in range(len(data)-1, 0, -1):
            if data[row_idx] and data[row_idx][0]:
                date_str = data[row_idx][0].strip()
                try:
                    return datetime.datetime.strptime(date_str, '%d-%m-%Y').date()
                except ValueError:
                    continue
        return None

    def save_daily_batch(self, fy_tab, month, day, daily_input):
        data = self._get_worksheet_data(fy_tab)
        worksheet = self.spreadsheet.worksheet(fy_tab)
        
        start_year = int(fy_tab.split('-')[0].replace("FY ", ""))
        year = start_year if month >= 4 else start_year + 1
        date_str = f"{day:02d}-{month:02d}-{year}"
        
        # Prepare the new row array (22 columns)
        new_row = ["" for _ in range(22)]
        new_row[0] = date_str
        
        for name, col_idx in CATEGORY_COLS.items():
            w = daily_input.get(name, {}).get('wagons', "0")
            if w == "":
                w = "0"
            new_row[col_idx] = str(w)
            
        row_idx = self._get_row_for_date(data, date_str)
        
        if row_idx == -1:
            # Find the first truly empty row
            sheet_row = len(data) + 1
            for i, row in enumerate(data):
                # Skip header row (0)
                if i > 0 and not any(str(cell).strip() for cell in row):
                    sheet_row = i + 1
                    break
            worksheet.update(values=[new_row], range_name=f"A{sheet_row}:V{sheet_row}")
        else:
            # Update existing row
            # gspread rowcol_to_a1 uses 1-indexed values
            sheet_row = row_idx + 1
            worksheet.update(values=[new_row], range_name=f"A{sheet_row}:V{sheet_row}")
            
        self._cache.pop(fy_tab, None)
        return True
