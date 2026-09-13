import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials as OAuthCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ['https://www.googleapis.com/auth/drive']

def get_drive_service():
    creds = None
    if os.path.exists('token.json'):
        creds = OAuthCredentials.from_authorized_user_file('token.json', SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('drive', 'v3', credentials=creds)
    return service

def find_or_create_folder(service, folder_name, parent_id=None, create_if_missing=True):
    """
    Finds a folder by name (and optional parent).
    If it doesn't exist and create_if_missing is True, creates it.
    Otherwise raises an exception.
    """
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
        
    results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    items = results.get('files', [])
    
    if not items:
        if not create_if_missing:
            raise Exception(f"Folder '{folder_name}' not found. Please ensure it exists and is shared with the service account with Editor access.")
            
        # Create folder
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parent_id:
            file_metadata['parents'] = [parent_id]
            
        folder = service.files().create(body=file_metadata, fields='id').execute()
        return folder.get('id')
    else:
        return items[0].get('id')

def upload_excel_to_drive(file_path, filename, fy_tab, month_name):
    """
    Uploads the generated Excel file to Google Drive.
    Structure: GM-MINI SHEETS / {fy_tab} / {month_name} / {filename}
    Returns the webViewLink of the uploaded file.
    """
    service = get_drive_service()
    
    # 1. Root Folder
    root_id = find_or_create_folder(service, 'GM-MINI SHEETS')
    
    # 2. FY Folder
    fy_id = find_or_create_folder(service, fy_tab, root_id)
    
    # 3. Month Folder
    month_id = find_or_create_folder(service, month_name, fy_id)
    
    # 4. Upload File
    file_metadata = {
        'name': filename,
        'parents': [month_id]
    }
    
    # Check if file with same name already exists to overwrite/update it, or just create new.
    query = f"name='{filename}' and '{month_id}' in parents and trashed=false"
    results = service.files().list(q=query, spaces='drive', fields='files(id)').execute()
    items = results.get('files', [])
    
    media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
    
    if items:
        # Update existing
        file_id = items[0].get('id')
        uploaded_file = service.files().update(
            fileId=file_id,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
    else:
        # Create new
        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
    return uploaded_file.get('webViewLink')
