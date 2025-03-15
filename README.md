# PDF Form to Excel Converter

A Python script that monitors a directory for new PDF files with form fields, extracts the data, and saves it to an Excel spreadsheet. Automatically deletes processed PDF files.

## Features

- 🕵️ Real-time directory monitoring for new PDF files
- 📊 Automatic Excel file creation with dynamic headers
- ✅ Preserves column order from first PDF's field structure
- 🧹 Auto-cleanup of processed PDF files
- 🛠 Error handling with detailed console logging

## Requirements

- Python 3.6+
- `watchdog` (file monitoring)
- `PyPDF2` (PDF processing)
- `openpyxl` (Excel file handling)

## Installation

1. Clone the repository:

```bash
   git clone https://github.com/yourusername/pdf-form-excel-converter.git
   cd pdf-form-excel-converter
```

2. Install dependencies:

```bash
    pip install watchdog PyPDF2 openpyxl
```
