import os
import re
import pandas as pd
from tqdm import tqdm
from pathlib import Path
from colorama import init, Fore, Style
import time
from openpyxl.worksheet.table import Table, TableStyleInfo

# Initialize colorama
init()

def print_ascii_art():
    """Print a cool ASCII art banner"""
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
    """
    Scan files in current directory for national numbers and create Excel report
    """
    print_ascii_art()
    time.sleep(0.5)  # Pause for dramatic effect
    
    print(Fore.CYAN + "=== Starting File Scanner ===" + Style.RESET_ALL)
    
    # Configuration
    pattern = r'(?<![0-9\u06F0-\u06F9])((?:[0-9\u06F0-\u06F9]{10})|(?:[0-9\u06F0-\u06F9]{9}))(?![0-9\u06F0-\u06F9])'
    root_dir = os.getcwd()
    print(Fore.CYAN + f"📂 Scanning directory: {root_dir}" + Style.RESET_ALL)
    
    matching_results = []
    
    # Gather all files
    print(Fore.YELLOW + "🔍 Gathering file list..." + Style.RESET_ALL)
    files = list(Path(root_dir).rglob('*'))
    
    total_files = len(files)
    print(Fore.YELLOW + f"📑 Scanning {total_files} file(s) for matching national codes..." + Style.RESET_ALL)
    
    # Scan files with progress bar
    for file in tqdm(files, desc="🔍 Scanning Files", bar_format="{l_bar}%s{bar}%s{r_bar}" % (Fore.GREEN, Style.RESET_ALL)):
        if file.is_file():  # Only process files, not directories
            matches = re.findall(pattern, file.name)
            for match in matches:
                national_number = match
                if len(national_number) == 9:
                    national_number = "0" + national_number
                matching_results.append({
                    'NationalNumber': national_number,
                    'FilePath': str(file.absolute())
                })
                print(Fore.GREEN + f"✅ Match: '{file.name}' -> National Number: {national_number}" + Style.RESET_ALL)
    
    print(Fore.CYAN + f"✨ Finished scanning. Found {len(matching_results)} matching file(s)." + Style.RESET_ALL)
    
    if not matching_results:
        print(Fore.RED + "❌ No matches found. Exiting..." + Style.RESET_ALL)
        return
    
    # Create DataFrames
    all_matches_df = pd.DataFrame(matching_results)
    unique_numbers_df = all_matches_df.drop_duplicates(subset=['NationalNumber'])
    
    print(Fore.GREEN + f"📊 Found {len(unique_numbers_df)} unique national number(s)." + Style.RESET_ALL)
    
    # Create Excel writer
    output_path = os.path.join(root_dir, "matching_files.xlsx")
    print(Fore.YELLOW + f"📝 Creating Excel file at: {output_path}" + Style.RESET_ALL)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Write sheets
        all_matches_df.to_excel(writer, sheet_name='All Matches', index=False)
        unique_numbers_df.to_excel(writer, sheet_name='Unique Numbers', index=False)
        
        # Get the workbook
        workbook = writer.book
        
        # Format both sheets
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            
            # Format as table
            data_range = f"A1:B{len(all_matches_df) + 1}" if sheet_name == 'All Matches' else f"A1:B{len(unique_numbers_df) + 1}"
            table = Table(displayName=f"Table_{sheet_name.replace(' ', '_')}", ref=data_range)
            
            # Add a default style
            style = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=True,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False
            )
            table.tableStyleInfo = style
            worksheet.add_table(table)
            
            # Set column widths
            worksheet.column_dimensions['A'].width = 20  # National Number column
            worksheet.column_dimensions['B'].width = 100  # File Path column
            
            # Format first column as text to preserve leading zeros
            for cell in worksheet['A']:
                cell.number_format = '@'
    
    print(Fore.CYAN + """
╔════════════════════════════════════╗
║ ✨ Excel file created successfully ✨ ║
╚════════════════════════════════════╝
""" + Style.RESET_ALL)

if __name__ == "__main__":
    scan_files()