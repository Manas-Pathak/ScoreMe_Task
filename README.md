# Detecting and Extracting Tables from PDFs
## [Project Report:](https://docs.google.com/document/d/1vJPrbR8V85enazd-zzrMQLVAqMFyE7TN2t6shS80P1s/edit?usp=sharing) 
## Overview
This project was developed as part of the **ScoreMe Hackathon Assignment** at **NIT Allahabad**. It focuses on detecting and extracting tables from system-generated PDFs, handling both regular and irregular shapes of tables, preserving the tabular data in **Excel files**. The solution adheres to the following constraints:
- **Does not use Tabula or Camelot**.
- **Does not involve converting PDFs into images**.

The tool uses Python and reliable libraries (`pdfplumber`, `pandas`, and `openpyxl`) to read PDF files, detect tabular structures based on text layout, and export structured data to Excel sheets.

---

## Features
- **Table Detection**: Extracts tables from PDFs (including tables with borders, no borders, and irregular layouts).
- **PDF Parsing**: Processes text-based PDFs using the `pdfplumber` library without requiring image conversions.
- **Error Handling**: Skips problematic pages with font encoding issues (e.g., "Non-Ascii85 digit found") while processing valid pages.
- **Excel Export**: Saves extracted data into Excel files with each worksheet corresponding to a PDF page.
- **Edge Case Handling**: Manages irregular tables, multilined cells, and empty or malformed pages gracefully.
- **Log Messages**: Provides clear logging for errors and skipped pages.

---

## Requirements
The project is implemented using Python and requires the following libraries:

- **Python Version**: 3.9+
- **Dependencies**:
  - `pdfplumber` (PDF text extraction)
  - `pandas` (Data manipulation for table extraction and Excel integration)
  - `openpyxl` (Excel file creation and formatting)
  
Install the required libraries via `pip`.

```bash
  pip install pdfplumber pandas openpyxl

    ---
