import json
import gspread
from google.oauth2.service_account import Credentials

try:
    with open('credentials.json', 'r') as f:
        data = json.load(f)
    
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(data, scopes=scope)
    print("Successfully loaded credentials info")
except Exception as e:
    print(f"Error: {e}")
