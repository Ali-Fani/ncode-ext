import os
from os import scandir
import re
import pylightxl as xl



ENABLE_LOGGING = True
SHOW_PROGRESS = True
EXCLUDE_DIRS = {'.git', '.venv', 'venv', 'env', '.env', 'ENV'}



def print_ascii_art():       
    if ENABLE_LOGGING:
        ascii_art = """
    ███████╗██╗██╗     ███████╗    ███████╗ ██████╗ █████╗ ███╗   ██╗
    ██╔════╝██║██║     ██╔════╝    ██╔════╝██╔════╝██╔══██╗████╗  ██║
    █████╗  ██║██║     █████╗      ███████╗██║     ███████║██╔██╗ ██║
    ██╔══╝  ██║██║     ██╔══╝      ╚════██║██║     ██╔══██║██║╚██╗██║
    ██║     ██║███████╗███████╗    ███████║╚██████╗██║  ██║██║ ╚████║
    ╚═╝     ╚═╝╚══════╝╚══════╝    ╚══════╝ ╚═════╝╚═╝  ╚═╝╚═╝  ╚═══╝
        """
        print(ascii_art)

def scan_files():
    print_ascii_art()
    print("=== Starting File Scanner ===")

    pattern = re.compile(r'(?<!\d)(\d{9}|\d{10})(?!\d)')
    root_dir = os.getcwd()
    print(f"Scanning directory: {root_dir}")

    matching_results = []
    total_files = 0

    def scan_directory(path):
        nonlocal total_files
        try:
            with scandir(path) as entries:
                for entry in entries:
                    if entry.is_dir() and entry.name not in EXCLUDE_DIRS:
                        scan_directory(entry.path)
                    elif entry.is_file():
                        total_files += 1
                        matches = pattern.findall(entry.name)
                        for match in matches:
                            national_number = match.zfill(10)
                            matching_results.append({
                                'NationalNumber': national_number,
                                'FilePath': entry.path
                            })
                            print(f"Match: '{entry.name}' -> National Number: {national_number}")
        except PermissionError:
            pass

    scan_directory(root_dir)
    print(f"Scanning {total_files} file(s) for matching national codes...")

    if not matching_results:
        print("No matches found. Exiting...")
        return

    unique_results = {entry['NationalNumber']: entry for entry in matching_results}.values()
    print(f"Found {len(unique_results)} unique national number(s).")

    data_all = [["NationalNumber", "FilePath"]] + [
        [entry['NationalNumber'], entry['FilePath']] for entry in matching_results
    ]
    data_unique = [["NationalNumber", "FilePath"]] + [
        [entry['NationalNumber'], entry['FilePath']] for entry in unique_results
    ]

    db = xl.Database()
    db.add_ws(ws="All Matches")
    for r, row in enumerate(data_all, start=1):
        for c, value in enumerate(row, start=1):
            db.ws("All Matches").update_index(row=r, col=c, val=value)

    db.add_ws(ws="Unique Numbers")
    for r, row in enumerate(data_unique, start=1):
        for c, value in enumerate(row, start=1):
            db.ws("Unique Numbers").update_index(row=r, col=c, val=value)

    output_path = os.path.join(root_dir, "matching_files.xlsx")
    xl.writexl(db, output_path)
    print("Excel file created successfully!")

if __name__ == "__main__":
    scan_files()
