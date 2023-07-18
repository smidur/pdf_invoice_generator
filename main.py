import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_num = filename.split("-")[0]
    deal_date = filename.split("-")[1]
    pdf.set_font(family="Times", size=20, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice #{invoice_num}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {deal_date}")
    pdf.output(f"PDF/{filename}.pdf")
