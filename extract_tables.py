import pdfplumber
import pandas as pd
import os
import warnings

# Suppress CropBox warnings
warnings.filterwarnings("ignore", category=UserWarning, module="pdfplumber")


def extract_table_from_pdf(pdf_path, output_folder):
    """
    Extract tables from a PDF file and save them to an Excel file.
    """
    with pdfplumber.open(pdf_path) as pdf:
        extracted_data = {}

        for page_number, page in enumerate(pdf.pages, start=1):
            try:
                # Extract text from the page
                page_text = page.extract_text()
                if not page_text:  # Skip pages with no text
                    print(f"Page {page_number}: No text found.")
                    continue

                print(f"Processing Page {page_number}...")
                table = extract_table(page_text)  # Helper function to process page text
                if table:  # Only add non-empty tables
                    extracted_data[f"Page {page_number}"] = table
                else:
                    print(f"Page {page_number}: No table data extracted.")
            except Exception as e:
                # Log detailed error for skipped pages
                print(f"Error processing Page {page_number} in '{pdf_path}': {e}")
                continue

    # Save the extracted data to Excel (only if data exists)
    save_to_excel(extracted_data, pdf_path, output_folder)


def extract_table(page_text):
    """
    Parse the page text and organize it into rows and columns.
    """
    lines = page_text.split("\n")  # Break page text into lines
    table_rows = []  # List to store all rows

    for line in lines:
        # Split each line into columns based on whitespace
        columns = [col.strip() for col in line.split(" ") if col.strip()]
        if columns:  # Only add non-empty rows
            table_rows.append(columns)

    return table_rows


def save_to_excel(data, pdf_path, output_folder):
    """
    Save extracted tables to an Excel file with each page as a worksheet.
    """
    # Check if there is any data to save
    if not data:
        print(f"No data extracted from {os.path.basename(pdf_path)}. Skipping Excel save.")
        return

    # Generate the Excel filename based on the PDF file name
    filename = os.path.basename(pdf_path).replace(".pdf", ".xlsx")
    excel_filepath = os.path.join(output_folder, filename)

    # Write the data to Excel
    with pd.ExcelWriter(excel_filepath, engine="openpyxl") as writer:
        for page_name, table_rows in data.items():
            # Convert rows into a Pandas DataFrame
            df = pd.DataFrame(table_rows)

            # Skip empty data
            if df.empty:
                print(f"Skipping empty data for {page_name}. No data extracted.")
                continue

            # Save each extracted table in a separate worksheet
            df.to_excel(writer, sheet_name=page_name, index=False, header=False)

    print(f"Saved extracted tables to: {excel_filepath}")


def batch_extract(input_folder, output_folder):
    """
    Extract tables from all PDF files in the input folder.
    """
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # List all PDF files in the input folder
    pdf_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith(".pdf")]
    if not pdf_files:
        print("No PDF files found in the input folder.")
        return

    for pdf_file in pdf_files:
        print(f"Processing: {pdf_file}")
        extract_table_from_pdf(pdf_file, output_folder)


if __name__ == "__main__":
    # Define input and output folders
    input_folder = "./input_pdfs"
    output_folder = "./output_excel"

    # Extract tables from PDFs in batch
    batch_extract(input_folder, output_folder)