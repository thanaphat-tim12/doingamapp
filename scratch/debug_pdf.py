import pdfplumber
import pandas as pd

def debug_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()
        df = pd.DataFrame(words)
        # Sort by top coordinate
        df = df.sort_values('top')
        print(f"--- Words in {file_path} ---")
        print(df[['text', 'x0', 'top', 'bottom']].to_string())

print("Analyzing template.pdf...")
try:
    debug_pdf("template.pdf")
except Exception as e:
    print(f"Error: {e}")
