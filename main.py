import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Read excel files in data frames
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Format name
    filename = Path(filepath).stem
    # invoice_number, invoice_date = filename.split("-")
    invoice_number = filename.split("-")[0]
    invoice_date = filename.split("-")[1]

    # Create PDF files
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_number}", ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
