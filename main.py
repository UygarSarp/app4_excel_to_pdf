from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="a4")
    pdf.add_page()
    filename = Path(filepath).stem
    number = filename.split("-")[0]
    date = filename.split("-")[1]
    pdf.set_font("Times", "B", 24)
    pdf.cell(w=10, h=12, txt=f"Invoice nr.{number}", ln=1)
    pdf.cell(w=10, h=12, txt=f"Date {date}", ln=1)
    pdf.output(f"PDFs/{filename}.pdf")