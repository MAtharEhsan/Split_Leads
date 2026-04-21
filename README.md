# Lead Splitter Tool

A Python script that reads leads from an Excel file, removes duplicates, splits them into multiple sheets with a specified number of leads per sheet, and logs all duplicates found.

## Features

- **Read from Excel:** Supports reading leads from all sheets in an Excel workbook
- **Remove Duplicates:** Automatically detects and removes duplicate leads (based on phone number or full row)
- **Split into Batches:** Divides leads into separate sheets with a user-specified batch size
- **Log Duplicates:** Prints all duplicate entries at the end with source sheet and row information
- **Interactive Input:** User-friendly prompts for file path and batch size

## Prerequisites

Before running the script, ensure you have the following installed:

### 1. Python 3.7+
Check if Python is installed:
```
python --version
```

If not installed, download from: https://www.python.org/downloads/

### 2. Required Python Package: openpyxl

Install the required package using pip:
```
pip install openpyxl
```

Or if you're using a virtual environment:
```
python -m pip install openpyxl
```

## How to Run

### Step 1: Run the Script
```
python split_leads.py
```

### Step 2: Follow the Prompts

The script will ask you two questions:

**1. Enter the path to the Excel file:**
```
Example: C:\Users\Home Laptop\Downloads\Python script\sample.xlsx
Or simply: sample.xlsx (if in the same directory)
```

**2. Enter the number of leads required per sheet:**
```
Example: 5000
(Must be a positive integer)
```

### Step 4: Review the Output

The script will display:
- Total rows loaded from the file
- Number of unique leads after deduplication
- Number of duplicates removed
- Number of sheets created
- List of all duplicate entries (if any)



## Example Usage

```
============================================================
LEAD SPLITTER TOOL
============================================================

This tool splits leads from an Excel file into multiple
sheets with a specified number of leads per sheet.

Enter the path to the Excel file: sample.xlsx

Enter the number of leads required per sheet (positive integer): 5000

Processing: sample.xlsx
Leads per sheet: 5000

Saved file: sample.xlsx
Total rows loaded: 26368
Unique leads: 26368
Duplicates removed: 0
Sheets created: 6

No duplicates found.
```



## Tips

For issues or questions, check the following:
1. Ensure Python and openpyxl are installed
2. Verify the Excel file path is correct
3. Ensure the Excel file is not corrupted
4. If permission error, Close the Excel file if it's open, then run the script again. Or wait a few moments and retry.
5. The number of leads per sheet must be a positive integer (e.g., 5000, not 5000.5 or -100).
6. Make sure the file has at least headers in the first row
