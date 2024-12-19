import os
import pandas as pd
from PyPDF2 import PdfReader
from tkinter import Tk, filedialog, messagebox
from pathlib import Path
from datetime import datetime

def format_date(date_str):
    """Format date string to [year].[month].[day].[hour:minute:second]."""
    try:
        # Handle PDF date format like D:YYYYMMDDHHmmSS+HH'mm'
        if date_str.startswith('D:'):
            date_str = date_str[2:]
            # Remove timezone information for simplicity
            date_str = date_str.split('+')[0].split('-')[0].split('Z')[0]
            date_obj = datetime.strptime(date_str, '%Y%m%d%H%M%S')
            return date_obj.strftime('%Y.%m.%d.%H:%M:%S')
        # Handle ISO 8601 format like YYYYMMDDHHmmSSZ
        elif date_str.endswith('Z') and len(date_str) == 15:
            date_str = date_str[:-1]  # Remove 'Z'
            date_obj = datetime.strptime(date_str, '%Y%m%d%H%M%S')
            return date_obj.strftime('%Y.%m.%d.%H:%M:%S')
        # Handle cases with additional Z00'00'
        elif date_str.endswith("Z00'00'"):
            date_str = date_str[:-6]  # Remove "Z00'00'"
            date_obj = datetime.strptime(date_str, '%Y%m%d%H%M%S')
            return date_obj.strftime('%Y.%m.%d.%H:%M:%S')
        # Handle Exiftool-like date format: YYYY:MM:DD HH:MM:SSZ
        elif ':' in date_str and ' ' in date_str:
            if date_str.endswith('Z'):
                date_str = date_str[:-1]  # Remove 'Z'
            date_obj = datetime.strptime(date_str, '%Y:%m:%d %H:%M:%S')
            return date_obj.strftime('%Y.%m.%d.%H:%M:%S')
        # Handle unexpected formats by logging them
        else:
            print(f"Unrecognized date format: {date_str}")
            return date_str  # Return original string if not in expected format
    except Exception as e:
        print(f"Error formatting date: {date_str}, Error: {e}")
        return date_str  # Return the original string if formatting fails

def extract_pdf_metadata(file_path):
    """Extract metadata from a PDF file using PyPDF2."""
    try:
        reader = PdfReader(file_path)
        metadata = reader.metadata
        formatted_metadata = {}

        if metadata:
            for key, value in metadata.items():
                if key in ('/CreationDate', '/ModDate') and value:
                    # Format creation and modification dates
                    formatted_metadata[key] = format_date(value)
                else:
                    formatted_metadata[key] = str(value)
        else:
            return {"Error": "No metadata found"}

        return formatted_metadata
    except Exception as e:
        return {"Error": str(e)}

def process_folder(folder_path):
    """Process all PDF files in the folder and extract their metadata."""
    data = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                metadata = extract_pdf_metadata(file_path)
                metadata["File Name"] = file
                data.append(metadata)
    return pd.DataFrame(data)

def save_to_excel(df, output_path):
    """Save the metadata DataFrame to an Excel file."""
    try:
        df.to_excel(output_path, index=False)
        print(f"Metadata saved to: {output_path}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")

def main():
    """Main function to select a folder, extract metadata, and save to Excel."""
    # Tkinter GUI for folder selection
    Tk().withdraw()  # Hide the root window

    # Display message to user
    messagebox.showinfo("Velg mappe", "Velg hvilken mappe du vil at jeg skal finne PDF'er og trekke ut metadata fra")

    folder_path = filedialog.askdirectory(title="Select a folder with PDF files")
    if not folder_path:
        print("No folder selected.")
        return

    print(f"Processing files in folder: {folder_path}")
    df = process_folder(folder_path)
    if df.empty:
        print("No PDF files found in the selected folder.")
    else:
        output_path = Path(folder_path) / "Metadata_PDF_Fra_Denne_Mappen.xlsx"
        save_to_excel(df, output_path)
        # Display success message
        messagebox.showinfo("Ferdig", "Metadata har blitt trukket ut fra alle PDF'ene i mappen din, og finnes nå i en Excel-fil i samme mappe.")

if __name__ == "__main__":
    main()
