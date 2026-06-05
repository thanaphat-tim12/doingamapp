import pandas as pd
from main import get_gspread_client, SHEET_URL, load_sheet_data

client = get_gspread_client()
sh = client.open_by_url(SHEET_URL)
sheet_names = [s.title for s in sh.worksheets()]
print("Sheet names:", sheet_names)

for sheet_name in sheet_names:
    print(f"\n--- Testing Sheet: {sheet_name} ---")
    ws = sh.worksheet(sheet_name)
    headers = ws.row_values(1)
    print("Raw headers:", headers)
    fields = [h.strip() for h in headers]
    print("Stripped headers:", fields)
    
    if 'ลำดับ' in fields:
        print("'ลำดับ' found in fields!")
    else:
        print("'ลำดับ' NOT found in fields!")

    full_df = load_sheet_data(sheet_name)
    print("full_df columns:", full_df.columns.tolist())
    
    if not full_df.empty and 'ลำดับ' in full_df.columns:
        seqs = pd.to_numeric(full_df['ลำดับ'], errors='coerce').dropna()
        if not seqs.empty:
            next_seq = str(int(seqs.max()) + 1)
            print("Calculated next_seq:", next_seq)
        else:
            print("seqs is empty after to_numeric")
    else:
        print("full_df is empty or 'ลำดับ' not in full_df.columns")
