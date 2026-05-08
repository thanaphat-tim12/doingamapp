import pdfplumber

pdf = pdfplumber.open(r'd:\Doingam\DoiNgamApp\ต่อใบอนุญาติ.pdf')
p = pdf.pages[0]
words = p.extract_words()

with open(r'd:\Doingam\DoiNgamApp\scratch\word_coords.txt', 'w', encoding='utf-8') as f:
    for w in words:
        # reportlab uses coordinates from bottom-left
        x = w['x0']
        y = p.height - w['bottom']
        f.write(f"{w['text']}: x={x:.1f}, y={y:.1f}\n")
