from pathlib import Path
from PyPDF2 import PdfReader
from reportlab.pdfgen import canvas

def pdf_to_ps(pdf_path, ps_path):
    pdf_reader = PdfReader(pdf_path)
    c = canvas.Canvas(ps_path)

    for page in pdf_reader.pages:
        text = page.extract_text()
        c.drawString(100, 750, text)

    c.save()

if __name__ == "__main__":
    input_pdf = Path("input.pdf")  # Replace with your PDF file path
    output_ps = Path("output.ps")  # Replace with your desired PostScript file path

    if input_pdf.exists():
        pdf_to_ps(input_pdf, output_ps)
        print(f"Converted {input_pdf} to {output_ps}")
    else:
        print(f"File {input_pdf} does not exist.")