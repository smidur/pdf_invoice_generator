import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_num, date = filename.split("-")
    # add invoice number
    pdf.set_font(family="Times", size=20, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice #{invoice_num}", ln=1)
    # add date
    pdf.set_font(family="Times", size=20, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=1)
    # add a space between the tables
    pdf.cell(w=0, h=10, ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]

    # add a header
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1, align="L")
    pdf.cell(w=60, h=8, txt=columns[1], border=1, align="L")
    pdf.cell(w=40, h=8, txt=columns[2], border=1, align="L")
    pdf.cell(w=30, h=8, txt=columns[3], border=1, align="L")
    pdf.cell(w=30, h=8, txt=columns[4], border=1, align="L", ln=1)

    # add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=60, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    # add total sum row
    total_sum = df['total_price'].sum()
    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=60, h=8, txt="")
    pdf.cell(w=40, h=8, txt="")
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=8, txt="Grand Total", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # add space between the tables
    pdf.cell(w=0, h=10, ln=1)

    # add total sum sentence
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=70, h=8, txt=f"The total price is {total_sum} EUR", ln=1)

    # add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=28, h=10, txt=f"PythonHow")
    pdf.image("pythonhow.png", h=8, w=8)

    pdf.output(f"PDF/{filename}.pdf")
