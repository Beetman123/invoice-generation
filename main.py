import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm",format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_num, invoice_date = filename.split("-")

    pdf.set_font(family="Times", size=16,style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice num.{invoice_num}", ln=1)

    pdf.set_font(family="Times", size=16,style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    column_names = df.columns
    column_names = [item.replace("_"," ").title() for item in column_names]

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=column_names[0], border=1)
    pdf.cell(w=50, h=8, txt=column_names[1], border=1)
    pdf.cell(w=40, h=8, txt=column_names[2], border=1)
    pdf.cell(w=30, h=8, txt=column_names[3], border=1)
    pdf.cell(w=30, h=8, txt=column_names[4], border=1, ln=1)

    for index,row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=f"{row["product_id"]}", border=1)
        pdf.cell(w=50, h=8, txt=f"{row["product_name"]}", border=1)
        pdf.cell(w=40, h=8, txt=f"{row["amount_purchased"]}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row["price_per_unit"]}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row["total_price"]}", border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # print total price
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    # Company name & logo
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
