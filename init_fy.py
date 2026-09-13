import gspread
from google.oauth2.service_account import Credentials
import sys

def init_financial_year(fy_name):
    # Setup connection
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file('credentials.json', scopes=scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open('GM-Mini-Database')
    
    # Try to add worksheet
    try:
        ws = spreadsheet.add_worksheet(title=fy_name, rows="500", cols="22")
        print(f"Created new tab: {fy_name}")
    except Exception as e:
        print(f"Error creating tab (might already exist): {e}")
        ws = spreadsheet.worksheet(fy_name)
        
    # Apply standard layout
    header = [
        'Date', 
        'CEMT', 'COAL', 'CONT', 'FERT', 'IMFT', 'POL', 'SALT', 'STEEL', 'DOC', 'FG', 'CHEM', 'AUTO', 'OTHERS',
        '', 
        'ADI', 'GIMB', 'BCT', 'BRC', 'RJT', 'BVP', 'RTM'
    ]
    
    ws.clear()
    # update A1:V1
    ws.update('A1:V1', [header])
    
    # Formatting
    requests = []
    sheet_id = ws.id
    
    # Bold Headers, Light Gray BG, Borders for A1:V1
    requests.append({
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': 0,
                'endRowIndex': 1,
                'startColumnIndex': 0,
                'endColumnIndex': 22
            },
            'cell': {
                'userEnteredFormat': {
                    'textFormat': {'bold': True},
                    'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment': 'MIDDLE',
                    'borders': {
                        'bottom': {'style': 'SOLID'},
                        'top': {'style': 'SOLID'},
                        'left': {'style': 'SOLID'},
                        'right': {'style': 'SOLID'}
                    }
                }
            },
            'fields': 'userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment,borders)'
        }
    })

    # Bold Date Column
    requests.append({
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': 1,
                'endRowIndex': 500,
                'startColumnIndex': 0,
                'endColumnIndex': 1
            },
            'cell': {
                'userEnteredFormat': {
                    'textFormat': {'bold': True},
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment': 'MIDDLE'
                }
            },
            'fields': 'userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)'
        }
    })

    # Center all
    requests.append({
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': 1,
                'endRowIndex': 500,
                'startColumnIndex': 1,
                'endColumnIndex': 22
            },
            'cell': {
                'userEnteredFormat': {
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment': 'MIDDLE'
                }
            },
            'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment)'
        }
    })

    # Freeze Row 1 and Col 1
    requests.append({
        'updateSheetProperties': {
            'properties': {
                'sheetId': sheet_id,
                'gridProperties': {
                    'frozenRowCount': 1,
                    'frozenColumnCount': 1
                }
            },
            'fields': 'gridProperties(frozenRowCount,frozenColumnCount)'
        }
    })

    # Resize Column O (Index 14)
    requests.append({
        'updateDimensionProperties': {
            'range': {
                'sheetId': sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': 14,
                'endIndex': 15
            },
            'properties': {
                'pixelSize': 30
            },
            'fields': 'pixelSize'
        }
    })

    spreadsheet.batch_update({'requests': requests})
    print(f"Successfully formatted {fy_name}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python init_fy.py 'FY 2027-28'")
    else:
        init_financial_year(sys.argv[1])
