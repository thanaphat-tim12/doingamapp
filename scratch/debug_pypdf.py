from pypdf import PdfReader

def get_info(path):
    reader = PdfReader(path)
    page = reader.pages[0]
    print(f"Box: {page.mediabox}")

get_info("template.pdf")
