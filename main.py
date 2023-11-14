# Import necessary libraries for file handling and data processing
import pandas as pd
import glob

# Get a list of file paths for all Excel files in the 'invoices' directory
filepaths = glob.glob("invoices/*xlsx")

# Read each Excel file into a Pandas DataFrame
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")