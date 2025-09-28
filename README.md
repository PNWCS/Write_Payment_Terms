# Payment Terms QuickBooks Import

A Python application with tkinter GUI for importing payment terms from Excel files to QuickBooks Desktop.

## Features

- **Payment Terms Import**: Import payment terms from Excel to QuickBooks Desktop
- **QuickBooks Integration**: Uses COM API (QBXML) to connect to QuickBooks Desktop
- **Simple GUI**: tkinter interface with import functionality
- **Excel Processing**: Read payment terms from specified Excel file format
- **File Selection**: Option to select and process other Excel files

## Installation

```bash
# Install dependencies
poetry install
```

## Usage

```bash
# Run the application
poetry run python -m xlsx_reader.main
```

## Excel File Format

The application reads payment terms from any Excel file you select:
- **Sheet name**: `payment_terms`
- **Column A (Name)**: Payment term name (e.g., "Net 30")
- **Column B (ID)**: Discount days (e.g., 30)

## Project Structure

```
xlsx_reader/
├── __init__.py
├── main.py              # Main entry point
├── gui.py               # tkinter GUI with payment terms import
└── excel_processor.py   # Excel processing and QuickBooks integration

tests/
├── __init__.py
└── test_excel_processor.py  # Tests for all functions
```

## QuickBooks Integration

- **COM API**: Uses QBXML via win32com for QuickBooks Desktop integration
- **Payment Methods**: Creates payment methods in QuickBooks
- **Error Handling**: Provides feedback on connection and import status

## Dependencies

- **openpyxl**: For reading Excel files
- **pywin32**: For QuickBooks COM API integration
- **tkinter**: For GUI (included with Python)
- **pytest**: For testing (dev dependency)

## Running Tests

```bash
poetry run pytest
```

Tests will create temporary Excel files automatically for testing your implementations.
