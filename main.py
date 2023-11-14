# Import necessary libraries for file handling and data processing
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Get a list of file paths for all Excel files in the 'invoices' directory
filepaths = glob.glob("invoices/*xlsx")

# Iterate through each file path in the 'filepaths' list
for filepath in filepaths:
    # Read data from the Excel file into a Pandas DataFrame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Create a new PDF document
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extract the invoice number from the filename
    filename = Path(filepath).stem
    invoice_number = filename.split("-")[0]

    # Add content to the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr. {invoice_number}")

    # Output the PDF with the filename derived from the original file
    pdf.output(f"PDFs/{filename}.pdf")