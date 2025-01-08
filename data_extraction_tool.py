import os
import sys
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import locale
import PyInstaller.__main__

# Define constants
OUTPUT_EXCEL = "output.xlsx"
OUTPUT_CSV = "output.csv"
SAMPLE_FILES = ["sample_invoice_1.pdf","sample_invoice_2.pdf"]

# Change directory to Current working directory for executable
if getattr(sys, 'frozen', False):  # Check if the script is frozen (i.e., running as a .exe)
    print('Running as executable')
    os.chdir(os.path.dirname(sys.executable))
else:
    print('Running as script')
    os.chdir(os.path.dirname(__file__))  # When running as a script

# Extracts the field from the content in tabular structure
def extract_value_data(content, field):
    amount = 0
    for table in content:
        for column in table:
            column = [item for item in column if item is not None]
            if field in column:
                # Remove any currency symbols and non-numeric characters 
                amount = re.sub(r'[^\d,\.]', '', column[1])
                amount = amount if ',' not in amount else amount.replace(",", ".")
    return float(amount)

# Extract invoice date which is both in tabular and text format
def extract_Invoice_date(content, field):
    if field != "Invoice date":
        locale.setlocale(locale.LC_ALL, 'de_DE.UTF-8')
        for table in content:
            for column in table:
                if field in column:
                    # Data is mapped since its vertically extracted
                    mapped_table = dict(zip(table[0], table[1]))
                    date = datetime.strptime(mapped_table['Date'], "%d. %B %Y").date()
                    return date.strftime("%d/%m/%Y")
    else:
        for line in content.split("\n"):
            if f'{field}:' in line:
                date = line.split(f'{field}: ')[1]
                date = datetime.strptime(date, "%b %d, %Y").date()
                return date.strftime("%d/%m/%Y")

# Creates the structure of the data extracted
def data_formatting(data, file_name, extracted_data, column):
    formatted_data = {}
    if data:
        if extracted_data:
            if file_name in extracted_data.values():
                formatted_data[column] = data
        else:
            formatted_data['File Name'] = file_name
            formatted_data[column] = data     
    return formatted_data

                

# Function to extract data from PDFs
def extract_data_from_pdf(file_name):
    extracted_data = {}
    try:
        with pdfplumber.open(file_name) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                # Data extraction has to do seperately by hardcoding filename since Total field is present in both sample files
                if file_name == 'sample_invoice_1.pdf':
                    amount = extract_value_data(tables, 'Gross Amount incl. VAT')
                    extracted_data.update(data_formatting(amount, file_name, extracted_data, 'Value'))
                    date = extract_Invoice_date(tables, 'Date')
                    extracted_data.update(data_formatting(date, file_name, extracted_data, 'Date'))

                elif file_name == "sample_invoice_2.pdf":
                    amount = extract_value_data(tables, 'Total')
                    extracted_data.update(data_formatting(amount, file_name, extracted_data, 'Value'))
                    # Date in sample_invoice_2 is not in tabular format, so perform text extraction
                    text = page.extract_text_simple()
                    date = extract_Invoice_date(text, 'Invoice date')
                    extracted_data.update(data_formatting(date, file_name, extracted_data, 'Date'))    
    except Exception as e:
        print(f"Error processing {file_name}: {e}")
    return extracted_data

# Write to Excel and CSV
def write_to_excel_and_csv(data):
    # Creates DataFrame and order the column names as File Name, Date and Value
    df = pd.DataFrame(data)[['File Name','Date', 'Value']]
    
    # Data frame is converted to CSV
    df.to_csv(OUTPUT_CSV, index=False, sep=";")

    # Data frame is converted to Excel (Sheet 1)
    with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
        df.to_excel(writer,index=False, sheet_name="Sheet 1")
        
        # Data frame is converted to Pivot table (Sheet 2)
        pivot_table = df.pivot_table(values="Value",
                index="Date",
                columns="File Name",
                aggfunc="sum",
                fill_value=0,
                margins=True,
                margins_name="Total")
        pivot_table.to_excel(writer, sheet_name="Sheet 2")

# Create executable
def create_executable():
    PyInstaller.__main__.run([
        "--name=DataExtractionTool",
        "--distpath=./",
        "--onefile",
        __file__,
    ])

if __name__ == "__main__":
    # Extract data from files
    extracted_data = []
    for file in SAMPLE_FILES:
        if os.path.exists(file):
            extracted_data.append(extract_data_from_pdf(file))
    
    # Write to Excel and CSV
    write_to_excel_and_csv(extracted_data)

    # Create executable ; Uncomment to create the executable file according to your OS
    # create_executable()

    print("Process completed. Files created:")
    print(f"- {OUTPUT_EXCEL}")
    print(f"- {OUTPUT_CSV}")