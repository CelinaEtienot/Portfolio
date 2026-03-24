import pdfplumber
import re
from collections import defaultdict
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import os

pdf_file = r"C:\Users\celin\Downloads\ML25113A022.pdf"
split_pdf_folder = r"C:\Users\celin\Downloads\GALL_SLR AMPs"
word_folder = r"C:\Users\celin\Downloads\GALL_SLR AMPs"

os.makedirs(split_pdf_folder, exist_ok=True)
os.makedirs(word_folder, exist_ok=True)

# -------- STEP 1: agrupar paginas por AMP --------

reader = PdfReader(pdf_file)
amp_pages = defaultdict(list)

start_collecting = False  # variable para ignorar páginas previas al primer AMP

for i, page in enumerate(reader.pages[20:], start=20):
    text = page.extract_text()

    if not text:
        continue

    match = re.search(r"(X\.E\d+)", text)

    if not match:
        continue

    amp_code = match.group(1)  # XI M24

    amp_pages[amp_code].append(i)

# -------- STEP 2: generar PDFs --------

amp_titles = {}

with pdfplumber.open(pdf_file) as pdf:
    for amp_code, pages in amp_pages.items():
        writer = PdfWriter()
        for p in pages:
            writer.add_page(reader.pages[p])

        first_page_text = pdf.pages[pages[0]].extract_text()
        first_line = first_page_text.split("\n")[0]
        title = first_line.strip()
        title = re.sub(r'[\\/*?:"<>|]', "", title)
        amp_titles[amp_code] = title

        filename = f"{title}.pdf"
        output_path = os.path.join(split_pdf_folder, filename)
        with open(output_path, "wb") as f:
            writer.write(f)

# -------- STEP 3: convertir cada PDF a Word --------

for file in os.listdir(split_pdf_folder):

    if not file.endswith(".pdf"):
        continue

    pdf_path = os.path.join(split_pdf_folder, file)

    title = file.replace(".pdf", "")
    word_path = os.path.join(word_folder, f"{title}.docx")

    cv = Converter(pdf_path)
    cv.convert(word_path)
    cv.close()

print("Conversión completa sin watermarks")