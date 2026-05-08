import io
import sys
import os

# Add the current directory to sys.path so we can import main
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from main import create_pdf_overlay
from pypdf import PdfReader, PdfWriter

context = {
    "p_license_book": "ทดสอบเล่มที่ 1",
    "p_license_no": "ทดสอบเลขที่ 2",
    "p_license_year": "2569",
    "p_name": "นายทดสอบ ทดสอบ",
    "p_nationality": "ไทย",
    "p_addr": "123",
    "p_moo": "1",
    "p_cid": "1234567890123",
    "p_phone": "0812345678",
    "p_shop": "ร้านทดสอบ",
    "p_type": "กิจการทดสอบ",
    "p_fee": "1000",
    "p_fee_text": "หนึ่งพันบาทถ้วน",
    "p_rcpt_book": "ร1",
    "p_rcpt_no": "ล2",
    "p_rcpt_date": "10 พฤษภาคม 2569",
    "issue_day": "10",
    "issue_month": "พฤษภาคม",
    "issue_year": "2569",
    "expire_day": "9",
    "expire_month": "พฤษภาคม",
    "expire_year": "2570",
}

# The create_pdf_overlay function uses "template.pdf"
# So let's back up template.pdf and copy template_v2.pdf to template.pdf
os.system('copy template.pdf template_bak.pdf')
os.system('copy template_v2.pdf template.pdf')

buffer = create_pdf_overlay(context)

with open(r'C:\Users\Tim\.gemini\antigravity\brain\a4b8be93-13a3-4803-95ae-aae5bb563ccf\scratch\test_output_v2.pdf', 'wb') as f:
    f.write(buffer.getvalue())

# Convert PDF to PNG to view it
import pdfplumber
pdf = pdfplumber.open(r'C:\Users\Tim\.gemini\antigravity\brain\a4b8be93-13a3-4803-95ae-aae5bb563ccf\scratch\test_output_v2.pdf')
page = pdf.pages[0]
im = page.to_image(resolution=150)
im.save(r'C:\Users\Tim\.gemini\antigravity\brain\a4b8be93-13a3-4803-95ae-aae5bb563ccf\scratch\test_output_v2.png')

# Restore template.pdf
os.system('copy template_bak.pdf template.pdf')
