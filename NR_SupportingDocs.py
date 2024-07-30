# This program transfers and appropriately names all supporting documents for all jobs completed under the NR contract, to be submitted for monthly claim.

import pandas as pd
import os
import shutil
import re
import sys
import warnings
from datetime import datetime
import Levenshtein
from concurrent.futures import ThreadPoolExecutor, as_completed

# Suppress specific warnings
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

def create_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

def extract_dates(date_string):
    dates = re.findall(r'\d{2}/\d{2}/\d{4}', date_string)
    return dates

def convert_date_format(date_string):
    converted_date = datetime.strptime(date_string, "%d/%m/%Y").strftime("%d.%m.%Y")
    return converted_date

def rename_file(original_path, new_folder_path, sheet_name):
    file_name = os.path.basename(original_path)
    base_name, ext = os.path.splitext(file_name)
    if "DKT" in file_name and "TC" in file_name:
        parts = base_name.split()
        new_name = f"TC DKT {parts[1].strip('-')}"
    else:
        parts = base_name.rsplit(' ', 2)
        new_name = f"{sheet_name} JOBSHEET {parts[-2]} {parts[-1]}"
    new_path = os.path.join(new_folder_path, new_name + ext)
    return new_path

def is_close_match(a, b, max_distance=2):
    if Levenshtein.distance(a, b) <= max_distance:
        return True
    return False


def search_and_copy_files(search_path, sheet_name, new_folder_path):    
    for root, _, files in os.walk(search_path):
        for file in files:
            if sheet_name in file:
                file_path = os.path.join(root, file)
                new_file_path = os.path.join(new_folder_path, file)
                shutil.copy2(file_path, new_file_path)


def process_sheet(excel_path, sheet_name, output_folder):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    try:
        value_B6 = df.iloc[4, 1].strip()
        value_B5 = df.iloc[3, 1].strip()
        new_folder_name = f"{sheet_name.strip()} {value_B6}, {value_B5}"
        new_folder_path = os.path.join(output_folder, new_folder_name)
        create_folder(new_folder_path)

        date_string = df.iloc[7, 1].strip()
        dates = extract_dates(date_string)
        converted_dates = [convert_date_format(date) for date in dates]

        for date in converted_dates:
            files_found = False
            month_name = datetime.strptime(date, "%d.%m.%Y").strftime("%B")
            year = datetime.strptime(date, "%d.%m.%Y").strftime("%Y")
            parent_folder = os.path.join(
                r"C:\Users\JulianNorman\Dropbox (Pipe Management Aus)\Operations\Scanned Timesheets\Sydney Water Contract",
                year, "NR"
            )

            for folder in os.listdir(parent_folder):
                if month_name in folder:
                    month_folder_path = os.path.join(parent_folder, folder)
                    if any(date in subfolder for subfolder in os.listdir(month_folder_path)):
                        date_folder_path = os.path.join(month_folder_path, date)

                        if os.path.exists(date_folder_path):
                            for root, _, files in os.walk(date_folder_path):
                                for file in files:
                                    if str(sheet_name).strip() in file:
                                        files_found = True
                                        file_path = os.path.join(root, file)
                                        if not os.path.exists(new_folder_path):
                                            os.makedirs(new_folder_path)
                                        new_file_path = rename_file(file_path, new_folder_path, sheet_name)
                                        shutil.copy2(file_path, new_file_path)
                                    else:
                                        filename_segments = file.split()
                                        for segment in filename_segments:
                                            if is_close_match(sheet_name.strip(), segment):
                                                print(f"WARNING -- POSSIBLE MATCH FOR {sheet_name} ON {date}: {file} (Segment: {segment})")
                                                break

            if not files_found:
                print(f"WARNING -- NO FILES FOUND FOR {sheet_name} ON {date}")

        # Search and copy files from the additional directory
        search_path = r"F:\NR CCI UPLOADS\PDF REPORTS"
        search_and_copy_files(search_path, sheet_name, new_folder_path)

    except KeyError as e:
        print(f"KeyError: {e} - Check if the cell references are correct in sheet {sheet_name}")
    except IndexError as e:
        print(f"IndexError: {e} - Check if the cell references are correct in sheet {sheet_name}")

def main(excel_path, output_folder):
    excel_data = pd.ExcelFile(excel_path)

    with ThreadPoolExecutor() as executor:
        futures = [executor.submit(process_sheet, excel_path, sheet_name, output_folder) for sheet_name in excel_data.sheet_names[2:]]
        for future in as_completed(futures):
            future.result()  # Ensure all threads complete

    print("\nPROGRAM FINISHED SUCCESSFULLY\n")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py <path_to_excel_sheet> <path_to_output_folder>")
        sys.exit(1)

    excel_path = sys.argv[1]
    output_folder = sys.argv[2]

    main(excel_path, output_folder)
