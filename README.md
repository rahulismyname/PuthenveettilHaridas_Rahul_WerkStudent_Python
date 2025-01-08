# Data Extraction Tool

The Data Extraction Tool is a simple application that extracts specific values from invoice PDFs, generates an Excel file with organized data and a pivot table, and creates a CSV file for further analysis.

---

## Features

1. **Data Extraction**:
   - Extracts values like "Gross Amount incl. VAT" or "Total" from invoice PDFs without considering the currency. (Since conversion of currency requires API)
   - Automatically detects and extracts dates from the invoices.

2. **Excel File Creation**:
   - Generates an Excel file (`output.xlsx`) with:
     - **Sheet 1**: Raw data containing columns for File Name, Date, and Value.
     - **Sheet 2**: A pivot table summarizing the date and value sum, and also by document name.

3. **CSV File Creation**:
   - Creates a CSV file (`output.csv`) with semicolon-separated values.

4. **Executable File**:
   - Provides a standalone executable file for windows (`.exe`) that runs the tool without any dependencies

---

## How to Use

### 1. Prerequisites
- Ensure you have the invoice PDF files to process.
- The `.exe` file must be in the same folder as the invoice PDFs.

### 2. Steps to Run

#### Option 1: Using the Executable File
1. Place your invoice PDF files in the same folder as the `DataExtractionTool.exe`.
2. Double-click on the `DataExtractionTool.exe` to run the application.
3. The tool will automatically:
   - Process the PDF files.
   - Generate `output.xlsx` and `output.csv` in the same folder.
4. Open the generated files to view the extracted data and analysis.

#### Option 2: Running the Python Script
1. Ensure Python is installed on your system.
2. Install the required libraries using:
   ```bash
   pip install -r requirements.txt
   ```
3. Place your invoice PDF files in the same folder as the script.
4. Run the script:
   ```bash
   python data_extraction_tool.py
   ```
5. The tool will generate `output.xlsx` and `output.csv` in the same folder.

---

## Output Files

1. **output.xlsx**:
   - **Sheet 1**: Contains extracted data in columns:
     - File Name
     - Date
     - Value
   - **Sheet 2**: A pivot table summarizing values by date and file name.

2. **output.csv**:
   - Contains the same data as Sheet 1 in the Excel file, but formatted as a CSV with semicolon (`;`) separators.

---

## Example Input and Output

### Input
- **PDF Files**: `sample_invoice_1.pdf`, `sample_invoice_2.pdf`.

### Output
- **Extracted Data (Sheet 1 in Excel)**:
  | File Name            | Date           | Value        |
  |----------------------|----------------|--------------|
  | sample_invoice_1.pdf | 01/03/2024     | 453.53       |
  | sample_invoice_2.pdf | 26/11/2016     | 950          |

- **Pivot Table (Sheet 2 in Excel)**:
  | Date           | sample_invoice_1.pdf  | sample_invoice_2.pdf  | Total     |
  |----------------|-----------------------|-----------------------|-----------|
  | 01/03/2024     | 45353                 | 0                     | 453.53    |
  | 26/11/2016     | 0                     | 950                   | 950       |
  | Total          | 45353                 | 950                   | 1403.53   |

---

## Troubleshooting

1. **No Output Generated**:
   - Check that the PDF files are in the same folder as the executable file or script.
   - Ensure the PDF files are not corrupted.

2. **Error Running the Executable**:
   - Ensure all files are in the same folder as the executable file.

---
