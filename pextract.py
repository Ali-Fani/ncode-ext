import os
import re
import pylightxl as xl
from colorama import init, Fore, Style
try:
    from tqdm import tqdm
except ImportError:
    tqdm = None

# Initialize colorama
init(autoreset=True)

# Flags to control logging and progress bar
ENABLE_LOGGING = False
SHOW_PROGRESS = False

# Directories to exclude
EXCLUDE_DIRS = {'.git', '.venv', 'venv', 'env', '.env', 'ENV'}

def log(message, color=Fore.WHITE):
    if ENABLE_LOGGING:
        print(color + message + Style.RESET_ALL)

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
        print(Fore.CYAN + ascii_art + Style.RESET_ALL)

def scan_files():
    print_ascii_art()
    log("=== Starting File Scanner ===", Fore.CYAN)

    # Precompile regex pattern for matching 9 or 10 digit national numbers
    pattern = r'(?<!\d)(\d{9}|\d{10})(?!\d)'
    regex = re.compile(pattern)

    root_dir = os.getcwd()
    log(f"Scanning directory: {root_dir}", Fore.CYAN)

    matching_results = []
    total_files = sum(len(files) for _, _, files in os.walk(root_dir))
    log(f"Scanning {total_files} file(s) for matching national codes...", Fore.YELLOW)

    use_progress = SHOW_PROGRESS and tqdm is not None
    pbar = tqdm(total=total_files, desc="Scanning Files") if use_progress else None

    for dirpath, dirnames, filenames in os.walk(root_dir):
        # Exclude specified directories
        dirnames[:] = [d for d in dirnames if d not in EXCLUDE_DIRS]

        for filename in filenames:
            if pbar:
                pbar.update(1)
            matches = regex.findall(filename)
            for match in matches:
                national_number = match.zfill(10)  # Ensure the number is 10 digits
                matching_results.append({
                    'NationalNumber': national_number,
                    'FilePath': os.path.join(dirpath, filename)
                })
                log(f"Match: '{filename}' -> National Number: {national_number}", Fore.GREEN)
    if pbar:
        pbar.close()

    log(f"Finished scanning. Found {len(matching_results)} matching file(s).", Fore.CYAN)

    if not matching_results:
        log("No matches found. Exiting...", Fore.RED)
        return

    # Deduplicate entries by national number
    unique_results = {entry['NationalNumber']: entry for entry in matching_results}.values()
    log(f"Found {len(unique_results)} unique national number(s).", Fore.GREEN)

    # Prepare data for worksheets
    data_all = [["NationalNumber", "FilePath"]] + [
        [entry['NationalNumber'], entry['FilePath']] for entry in matching_results
    ]
    data_unique = [["NationalNumber", "FilePath"]] + [
        [entry['NationalNumber'], entry['FilePath']] for entry in unique_results
    ]

    # Create a new pylightxl Database and add worksheets
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
    log("Excel file created successfully!", Fore.CYAN)

if __name__ == "__main__":
    scan_files()
