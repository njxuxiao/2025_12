import pandas as pd
import os
import sys
from zipfile import BadZipFile

def convert_excel_to_txt(excel_filepath):
    """
    Reads an Excel file and converts each sheet into a tab-delimited TXT file.

    Args:
        excel_filepath (str): The path to the .xlsx file.
    """
    if not os.path.exists(excel_filepath):
        print(f"Error: File not found '{excel_filepath}'. Skipping.")
        return

    if not excel_filepath.lower().endswith('.xlsx'):
        print(f"Error: File '{excel_filepath}' is not an .xlsx file. Skipping.")
        return

    print(f"\n--- Processing '{excel_filepath}' ---")
    
    try:
        xls = pd.ExcelFile(excel_filepath, engine='openpyxl')
        sheet_names = xls.sheet_names

        if not sheet_names:
            print("  No sheets found in this file.")
            return

        base_filename = os.path.splitext(excel_filepath)[0]

        for sheet_name in sheet_names:
            print(f"  -> Converting sheet: '{sheet_name}'...")
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Clean up column names just in case
            df.columns = df.columns.str.strip()

            # Define the output txt filename
            output_filename = f"{base_filename}_{sheet_name}.txt"
            
            # Save to a tab-delimited txt file
            df.to_csv(output_filename, sep='\t', index=False, encoding='utf-8-sig')
            print(f"     Success! Saved to '{output_filename}'")
            
    except BadZipFile:
        print(f"  Error: The file '{excel_filepath}' seems to be corrupted.")
        print("  [SOLUTION]: Please open it in Excel, 'Save As' -> 'Excel Workbook (*.xlsx)' to fix it.")
    except Exception as e:
        print(f"  An unexpected error occurred: {e}")

if __name__ == '__main__':
    # --- HOW TO USE ---
    # 1. Place this script in the same folder as your Excel files.
    # 2. Add the names of all the Excel files you want to convert into the list below.
    files_to_convert = [
        "1_21_07.xlsx", 
        "2_22_12.xlsx", 
        "3_22_07.xlsx", 
        "4_21_12.xlsx",
        "5_20_12.xlsx", 
        "6_19_12.xlsx",
        "7_19_07.xlsx",
        "all.xlsx"
    ]

    print("Starting batch conversion from Excel (.xlsx) to Text (.txt)...")

    for file in files_to_convert:
        convert_excel_to_txt(file)
    
    print("\nBatch conversion finished!")
