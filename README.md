# Excel Column Extractor

A Python utility that extracts the first three columns from an Excel file and saves them to a new Excel file.

## Features

- Reads any Excel file format supported by pandas
- Extracts only the first three columns
- Creates a new Excel file with the extracted data
- Provides detailed logging
- Handles errors gracefully

## Requirements

- Python 3.6+
- pandas
- openpyxl

## Installation

1. Clone this repository or download the source code
2. Install required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Command Line

```bash
python excel_processor.py <input_excel_file> [output_excel_file]
```

Where:
- `<input_excel_file>`: Path to the Excel file you want to process (required)
- `[output_excel_file]`: Path where the output file will be saved (optional)
  - If not provided, the output will be saved with the same name as the input file but with "_processed" appended

### Python API

You can also use the script as a module in your Python code:

```python
from excel_processor import process_excel_file

result = process_excel_file('input.xlsx', 'output.xlsx')
if result:
    print(f"Successfully processed file: {result}")
```

## Error Handling

The script includes comprehensive error handling:
- Validates that the input file exists
- Ensures the input file has at least three columns
- Logs detailed error information if any issues occur

## Last Updated

May 3, 2025