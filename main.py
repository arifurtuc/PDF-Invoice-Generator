# Import necessary libraries for file handling and data processing
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Get a list of file paths for all Excel files in the 'invoices' directory
filepaths = glob.glob("invoices/*.xlsx")

# Iterate through each file path in the 'filepaths' list
for filepath in filepaths:
    # Create a new PDF document
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extract the invoice number, date from the filename
    filename = Path(filepath).stem
    invoice_number, date = filename.split("-")

    # Add invoice nr. to the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr. {invoice_number}", ln=1)

    # Add date to the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Read data from the Excel file into a Pandas DataFrame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Convert column names to a more readable format
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]

    # Set font, size, and text color for the table headers
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)

    # Create table headers in the PDF
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=65, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Iterate over DataFrame rows and populate the PDF with data
    for index, row in df.iterrows():
        # Set font and text color for the table data
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)

        # Add data to the PDF table cells
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=65, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Output the PDF with the filename derived from the original file
    pdf.output(f"PDFs/{filename}.pdf")
