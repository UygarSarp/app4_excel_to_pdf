from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="a4")
    pdf.add_page()

    filename = Path(filepath).stem
    number, date = filename.split("-")

    pdf.set_font("Times", "B", 24)
    pdf.cell(w=10, h=12, txt=f"Invoice nr.{number}", ln=1)
    pdf.cell(w=10, h=12, txt=f"Date {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    liste = df.columns
    liste = [item.replace("_", " ").title() for item in liste]
    pdf.set_font("Times", "B", 12)
    pdf.cell(w=25, h=8, txt=liste[0], border=1, align="L")
    pdf.cell(w=60, h=8, txt=liste[1], border=1, align="L")
    pdf.cell(w=40, h=8, txt=liste[2], border=1, align="L")
    pdf.cell(w=40, h=8, txt=liste[3], border=1, align="L")
    pdf.cell(w=25, h=8, txt=liste[4], border=1, ln=1, align="L")

    pdf.set_font("Times", "I", 10)
    pdf.set_text_color(80, 80, 80)
    # total = 0
    for i, row in df.iterrows():
        pdf.cell(w=25, h=8, txt=str(row["product_id"]), border=1, align="L")
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1, align="L")
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1, align="L")
        pdf.cell(w=40, h=8, txt=str(row["price_per_unit"]), border=1, align="L")
        pdf.cell(w=25, h=8, txt=str(row["total_price"]), border=1, ln=1, align="L")
        # total = total + row["total_price"]

    total = df["total_price"].sum()
    pdf.cell(w=25, h=8, border=1, align="L")
    pdf.cell(w=60, h=8, border=1, align="L")
    pdf.cell(w=40, h=8, border=1, align="L")
    pdf.cell(w=40, h=8, border=1, align="L")
    pdf.cell(w=25, h=8, border=1, txt=str(total), ln=1, align="L")

    pdf.set_font("Times", "B", 24)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=40, h=8, align="L", ln=1)
    pdf.cell(w=10, h=12, txt=f"The total due amount is {total} Euros.", ln=1)
    pdf.cell(w=10, h=12, txt="The Rock")
    pdf.image("stone.png", w=30, x=50)

    pdf.output(f"PDFs/{filename}.pdf")
