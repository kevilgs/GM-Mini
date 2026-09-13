import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import datetime

# Define constants for row mappings (0-indexed in python list)
COMMODITY_ROWS = {
    'CEMT': 2, 'COAL': 3, 'CONT': 4, 'FERT': 5, 'IMFT': 6, 'POL': 7, 
    'SALT': 8, 'STEEL': 9, 'DOC': 10, 'FG': 11, 'CHEM': 12, 'AUTO': 13, 'OTHERS': 14
}

DIVISION_ROWS = {
    'ADI': 18, 'GIMB': 19, 'BCT': 20, 'BRC': 21, 'RJT': 22, 'BVP': 23, 'RTM': 24
}

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
        """Fetches the entire worksheet as a 2D list with caching"""
        if fy_tab not in self._cache or bypass_cache:
            worksheet = self.spreadsheet.worksheet(fy_tab)
            self._cache[fy_tab] = worksheet.get_all_values()
        return self._cache[fy_tab]
        
    def _get_year_for_month(self, fy_tab: str, month: int) -> int:
        start_year = int(fy_tab.split('-')[0].replace("FY ", ""))
        return start_year if month >= 4 else start_year + 1
        
    def get_wagon_count(self, fy_tab, category, name, month, day):
        """
        Gets the wagon count for a specific date from the 419-column grid.
        Category: 'commodity' or 'division'
        Name: e.g. 'CEMT' or 'ADI'
        Month: 4 (April) to 3 (March)
        Day: 1 to 31
        """
        data = self._get_worksheet_data(fy_tab)
        
        row_idx = COMMODITY_ROWS[name] if category == 'commodity' else DIVISION_ROWS[name]
        
        # The layout repeats month blocks.
        # Format for a month block:
        # Col N: 'YYYY-MM-01' (empty data)
        # Col N+1: 'Head'
        # Col N+2: 'YYYY-MM-01' (Day 1 data)
        # Col N+3: 'YYYY-MM-02' (Day 2 data)
        # Therefore, for Day D, the column is (Col N+1) + D
        
        year = self._get_year_for_month(fy_tab, month)
        month_str = f"{year}-{month:02d}-01"
        
        head_col = -1
        for i in range(len(data[0]) - 1):
            if data[0][i] == month_str and data[0][i+1] == "Head":
                head_col = i + 1
                break
                
        if head_col == -1:
            return 0.0
            
        target_col = head_col + day
        
        val = data[row_idx][target_col]
        try:
            return float(val) if val != "" else 0.0
        except ValueError:
            return 0.0
            
    def get_month_sum(self, fy_tab, category, name, month, up_to_day=None):
        """
        Calculates the sum of wagons for a given month.
        If up_to_day is provided, it calculates MTD (Month to Date).
        """
        data = self._get_worksheet_data(fy_tab)
        row_idx = COMMODITY_ROWS[name] if category == 'commodity' else DIVISION_ROWS[name]
        
        year = self._get_year_for_month(fy_tab, month)
        month_str = f"{year}-{month:02d}-01"
        
        head_col = -1
        for i in range(len(data[0]) - 1):
            if data[0][i] == month_str and data[0][i+1] == "Head":
                head_col = i + 1
                break
                
        if head_col == -1:
            return 0.0
            
        max_days = MONTH_DAYS[month]
        end_day = up_to_day if up_to_day is not None else max_days
        
        total = 0.0
        for d in range(1, end_day + 1):
            col = head_col + d
            val = data[row_idx][col]
            try:
                total += float(val) if val != "" else 0.0
            except ValueError:
                pass
                
        return total
        
    def get_fy_progressive_sum(self, fy_tab, category, name, current_month, current_day):
        """
        Calculates total wagons from April 1st up to the given month and day.
        """
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
        """
        Scans the database columns from right to left to find the latest date
        that has non-empty values in the data rows.
        """
        data = self._get_worksheet_data(fy_tab)
        
        # Iterate columns backwards, skipping the last column (which is TOTAL usually)
        for col in range(len(data[0]) - 1, 1, -1):
            header = data[0][col].strip()
            # Check if header is a valid date string YYYY-MM-DD
            if header and len(header) == 10 and header.count('-') == 2:
                # Ensure it's not a month block header (which has 'Head' in the next col)
                if col < len(data[0]) - 1 and data[0][col+1] == "Head":
                    continue
                
                # Check if any data row has values
                has_data = False
                for r in list(range(2, 15)) + list(range(18, 25)):
                    if r < len(data) and col < len(data[r]) and data[r][col].strip() != "":
                        has_data = True
                        break
                
                if has_data:
                    return datetime.datetime.strptime(header, "%Y-%m-%d").date()
                    
        return None

    def save_daily_batch(self, fy_tab, month, day, daily_input):
        """
        daily_input: dict mapping name to {'wagons': W, 'rakes': R}
        Updates all commodities (rows 2-14) and divisions (rows 18-24) in 2 API calls.
        """
        data = self._get_worksheet_data(fy_tab)
        worksheet = self.spreadsheet.worksheet(fy_tab)
        
        year = self._get_year_for_month(fy_tab, month)
        month_str = f"{year}-{month:02d}-01"
        
        head_col = -1
        for i in range(len(data[0]) - 1):
            if data[0][i] == month_str and data[0][i+1] == "Head":
                head_col = i + 1
                break
                
        if head_col == -1:
            print(f"Error: Could not find column for month {month}")
            return False
            
        target_col = head_col + day
        col_letter = gspread.utils.rowcol_to_a1(1, target_col + 1)[0:-1]
        
        # Prepare commodity column vector (Rows 3 to 15, index 2 to 14)
        comm_vals = []
        for name in COMMODITY_ROWS.keys():
            w = daily_input.get(name, {}).get('wagons', "")
            comm_vals.append([w])
            
        # Prepare division column vector (Rows 19 to 25, index 18 to 24)
        div_vals = []
        for name in DIVISION_ROWS.keys():
            w = daily_input.get(name, {}).get('wagons', "")
            div_vals.append([w])
            
        # Update commodities
        worksheet.update(f"{col_letter}3:{col_letter}15", comm_vals)
        # Update divisions
        worksheet.update(f"{col_letter}19:{col_letter}25", div_vals)
        
        self._cache.pop(fy_tab, None)
        return True

if __name__ == "__main__":
    db = GSheetsDB()
    # Quick Test
    print("Testing read CEMT for August 1, 2025...")
    val = db.get_wagon_count("FY 2025-26", "commodity", "CEMT", 8, 1)
    print(f"August 1 CEMT: {val}")
    
    print("Testing MTD sum for CEMT up to August 5...")
    mtd = db.get_month_sum("FY 2025-26", "commodity", "CEMT", 8, 5)
    print(f"August 1-5 MTD CEMT: {mtd}")
