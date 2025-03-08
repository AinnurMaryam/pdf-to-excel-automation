# PDF to Excel Data Automation Script
This Python script automates data extraction from PDF reports and updates an Excel tracking file, reducing manual effort and improving data accuracy.

## Features
✅ Extracts structured data from PDFs using pdfplumber
✅ Cleans and standardizes numeric values (handles percentages, commas, and spaces)
✅ Dynamically updates an Excel file with extracted data using openpyxl
✅ Supports monthly data tracking with automatic column selection
✅ Converts extracted text values into numerical format

## Installation

### Prerequisites
Ensure you have Python installed (>=3.7). Install the required dependencies:
  ```bash
  pip install pdfplumber openpyxl
  ```

## Usage
1️⃣ Modify the script to specify your PDF and Excel file paths:
  ```python
  pdf_path = "path/to/your/file.pdf"
  excel_path = "path/to/your/file.xlsx"
  ```
2️⃣ Run the script:
  ```bash
  python script.py
  ```
3️⃣ The script will extract data from the PDF, clean it, and update the corresponding Excel sheet.

## Customization
Update the table row/column mappings based on your PDF structure.
Modify get_cell_for_month() if your Excel file has different column mappings.

## Notes
🔹 This script is generic and does not include specific company data.
🔹 If your PDF structure differs, adjust the table indexing in extract_table().

## License
This project is open-source under the MIT License.
