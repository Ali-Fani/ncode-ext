## National Number Extractor
This script extracts national numbers from current directory and subdirectories.
It saves the results in an Excel file.

### Build
```
nuitka --onefile --windows-icon-from-ico=E:\projects\ncode-ext\icon.ico --product-name='National Number Extractor' --product-version=1.0.0.0 --file-description='Extracts names with national number and puts them in excel file' --copyright='Ali Fani' --report=compilation-report.xml .\pextract.py
```

### Usage
```
python .\pextract.py
```
or
```
.\petract.exe
```
