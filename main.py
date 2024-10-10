import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=16)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(["amount_purchased"]), border=1)
        pdf.cell(w=70, h=8, txt=str(["price_per_unit"]), border=1)
        pdf.cell(w=70, h=8, txt=str(["total_price"]), border=1, ln=1)

    pdf.output(f"~/PycharmProjects/Excel_to_pdf/PDFs/{filename}.pdf")
