import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
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

    # Read excel files in data frames
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add column headers
    column_headers = df.columns
    column_headers = [item.replace("_", " ").title() for item in column_headers]

    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(50, 50, 50)
    pdf.cell(w=30, h=8, txt=column_headers[0], border=1)
    pdf.cell(w=70, h=8, txt=column_headers[1], border=1)
    pdf.cell(w=35, h=8, txt=column_headers[2], border=1)
    pdf.cell(w=30, h=8, txt=column_headers[3], border=1)
    pdf.cell(w=30, h=8, txt=column_headers[4], border=1, ln=1)

    # Populate table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(70, 70, 70)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=35, h=8, txt=str(row['amount_purchased']), align="R", border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), align="R", border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), align="R", border=1, ln=1)

    # Total price row
    total_price = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_fill_color(230)
    pdf.cell(w=30, h=8, fill=True, txt="", border=1)
    pdf.cell(w=70, h=8, fill=True, txt="", border=1)
    pdf.cell(w=35, h=8, fill=True, txt="", border=1)
    pdf.cell(w=30, h=8, fill=True, txt="", border=1)
    pdf.cell(w=30, h=8, fill=True, txt=str(total_price), border=1, align="R", ln=1)

    # Add total sum in writing
    pdf.set_font(family="Times", style="B", size=14)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(w=50, h=8, txt=f"The total price is {total_price}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=50, h=8, txt=f"Magnolia Apartments")
    pdf.image("magnolia-bitmap.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
