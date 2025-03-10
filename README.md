## National Number Extractor
This script extracts national numbers from current directory and subdirectories.
It saves the results in an Excel file.

# Install requirements
```pip install -r requirements.txt```
### old build with nuitka
```
nuitka --onefile --windows-icon-from-ico=./icon.ico --product-name='National Number Extractor' --product-version=1.0.0.0 --file-description='Extracts names with national number and puts them in excel file' --copyright='Ali Fani' --report=compilation-report.xml .\pextract.py 
```
### Build with pyinstaller
```
pyinstaller --onefile --icon=./icon.ico --name='National Number Extractor' --version-file=./version.txt .\pextract.py
```
### Usage
```
python .\pextract.py
```
or
```
.\petract.exe
```
