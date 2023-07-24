from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

# = everything ending in xlsx
filepaths = glob.glob("Invoices/*xlsx")


for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # headers, makes the filename string a dynamic intelligent filepath
    # so can extract the stem: not the folder or extension
    filename = Path(filepath).stem
    invoice_number, date = filename.split("-")

    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=50, h=12, txt=f"Invoice nr. {invoice_number}", align="L", ln=1)

    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=50, h=12, txt=f"Date {date}", align="L", ln=1)

    pdf.ln(3)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # table with invoice info
    # set table headers

    df_columns = df.columns
    columns_header = [item.replace("_", " ").title() for item in df.columns]

    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=12, txt=columns_header[0], border=1)
    pdf.cell(w=70, h=12, txt=columns_header[1], border=1)
    pdf.cell(w=40, h=12, txt=columns_header[2], border=1)
    pdf.cell(w=30, h=12, txt=columns_header[3], border=1)
    pdf.cell(w=30, h=12, txt=columns_header[4], border=1, ln=1)

    # add row
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=12, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=12, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=12, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=12, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=12, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/Invoice_{filename}.pdf")
