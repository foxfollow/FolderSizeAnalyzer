import os
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import sys


def get_folder_size(folder_path):
    total_size = 0
    for path, _, files in os.walk(folder_path):
        for f in files:
            fp = os.path.join(path, f)
            try:
                total_size += os.path.getsize(fp)
            except Exception as e:
                create_error_log(f"Error getting file size: {str(e)}")
                continue  # Log the error and continue scanning
    return total_size


def convert_size(size_in_bytes):
    # Convert size to KB, MB, or GB
    if size_in_bytes < 1024:
        return size_in_bytes, 'Bytes'
    elif size_in_bytes < 1024 * 1024:
        return round(size_in_bytes / 1024, 2), 'KB'
    elif size_in_bytes < 1024 * 1024 * 1024:
        return round(size_in_bytes / (1024 * 1024), 2), 'MB'
    else:
        return round(size_in_bytes / (1024 * 1024 * 1024), 2), 'GB'


def scan_folders(folder_path):
    # Create a list of all folders
    all_folders = [x[0] for x in os.walk(folder_path)]
    total_folders = len(all_folders)
    folder_sizes = []

    for i, folder in enumerate(all_folders):
        # Skip folders...
        # if "Windows" in folder or "OneDrive" in folder or "OneDrive - SCPC of 3SCIP" in folder:
        #     continue

        try:
            folder_size = get_folder_size(folder)
            size, unit = convert_size(folder_size)
            folder_sizes.append((folder, float(size), unit))  # Convert size to float

            # Calculate progress
            progress_percent = ((i + 1) / total_folders) * 100
            print(f"Scanning folder {folder}... ({progress_percent:.2f}% complete)")
        except Exception as e:
            create_error_log(f"Error scanning folder: {str(e)}")
            continue  # Log the error and continue scanning

    # Define the order of units
    unit_order = {'GB': 1, 'MB': 2, 'KB': 3, 'Bytes': 4}

    # Sort by unit first, then by size
    folder_sizes.sort(key=lambda x: (unit_order[x[2]], -x[1]))  # Sort sizes in descending order

    return folder_sizes


def create_excel(folder_sizes):
    df = pd.DataFrame(folder_sizes, columns=['Folder Path', 'Size', 'Unit'])
    now = datetime.now()
    date_time = now.strftime("%m-%d-%H-%M")
    excel_path = os.path.join(folder_path, f'outp{date_time}.xlsx')
    df.to_excel(excel_path, index=False)


def create_excel_xl(folder_sizes):
    wb = Workbook()
    ws = wb.active
    ws.append(['Folder Path', 'Size', 'Unit'])
    for folder_size in folder_sizes:
        try:
            ws.append(folder_size)
        except Exception as e:
            create_error_log(f"Error appending folder size: {str(e)}")
    now = datetime.now()
    date_time = now.strftime("%m-%d-%H-%M")
    excel_path = os.path.join(os.getcwd(), f'outp{date_time}.xlsx')
    wb.save(excel_path)


def create_error_log(error_name):
    now = datetime.now()
    date_time = now.strftime("%m-%d-%H")
    date_time2 = now.strftime("%m-%d-%H-%M-%S")
    log_path = os.path.join(os.getcwd(), f'error_log_{date_time}.txt')  # Use os.getcwd() to get the current working directory
    with open(log_path, 'a') as f:
        f.write(f"{date_time2}: {error_name}\n")


folder_path = sys.argv[1] if len(sys.argv) > 1 else input("Enter the folder path: ")
print("Scanning folders...")
try:
    folder_sizes = scan_folders(folder_path)
    print("Creating Excel file...")
    create_excel_xl(folder_sizes)
    print("Excel file created successfully.")
except Exception as e:
    create_error_log(str(e))
    print("An error occurred. Please check the error log.")
    input("Press Enter to exit...")

input("Press Enter to exit...")


