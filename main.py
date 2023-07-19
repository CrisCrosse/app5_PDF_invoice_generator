from fpdf import FPDF
import pandas as pd
import glob

# = everything ending in xlsx
filepaths = glob.glob("Invoices/*xlsx")


for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
