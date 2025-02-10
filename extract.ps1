@echo off
# -------------------------
# CONFIGURATION & INITIAL LOGGING
# -------------------------
$pattern = '(?<![0-9\u06F0-\u06F9])((?:[0-9\u06F0-\u06F9]{10})|(?:[0-9\u06F0-\u06F9]{9}))(?![0-9\u06F0-\u06F9])'
Write-Host "=== Starting File Scan Script ===" -ForegroundColor Cyan

$rootDir = Get-Location
Write-Host "Scanning directory: $rootDir" -ForegroundColor Cyan

$matchingResults = @()

# -------------------------
# SCAN FOR FILES WITH PROGRESS BAR
# -------------------------
Write-Host "Gathering file list..." -ForegroundColor Yellow
$files = Get-ChildItem -Path $rootDir -Recurse -File
$totalFiles = $files.Count
$i = 0

Write-Host "Scanning $totalFiles file(s) for matching national codes..." -ForegroundColor Yellow
foreach ($file in $files) {
    $i++
    Write-Progress -Activity "Scanning Files" -Status "Processing file $i of $totalFiles" -PercentComplete (($i / $totalFiles) * 100)
    if ($file.Name -match $pattern) {
        $nationalNumber = $Matches[1]
        if ($nationalNumber.Length -eq 9) {
            $nationalNumber = "0" + $nationalNumber
        }
        $matchingResults += [PSCustomObject]@{
            NationalNumber = $nationalNumber
            FilePath       = $file.FullName
        }
        Write-Host "Match: '$($file.Name)' -> National Number: $nationalNumber" -ForegroundColor Green
    }
}
Write-Host "Finished scanning. Found $($matchingResults.Count) matching file(s)." -ForegroundColor Cyan

# -------------------------
# CREATE EXCEL AND WRITE "ALL MATCHES" SHEET
# -------------------------
Write-Host "Launching Excel and creating workbook..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()

# Assume the new workbook has at least one sheet; rename first sheet.
$worksheet1 = $workbook.Worksheets.Item(1)
$worksheet1.Name = "All Matches"

# Set Column A to Text format (to preserve any leading zero).
$worksheet1.Columns.Item(1).NumberFormat = "@"

$row = 1
$worksheet1.Cells.Item($row, 1).Value2 = "National Number"
$worksheet1.Cells.Item($row, 2).Value2 = "File Path"
$row++

$totalMatches = $matchingResults.Count
$i = 0
Write-Host "Writing data to 'All Matches' sheet..." -ForegroundColor Yellow
foreach ($result in $matchingResults) {
    $i++
    Write-Progress -Activity "Writing to All Matches sheet" -Status "Row $i of $totalMatches" -PercentComplete (($i / $totalMatches) * 100)
    $worksheet1.Cells.Item($row, 1).Value2 = $result.NationalNumber
    $worksheet1.Cells.Item($row, 2).Value2 = $result.FilePath
    $row++
}

$lastRow = $row - 1
$range1 = $worksheet1.Range("A1:B$lastRow")
Write-Host "Converting 'All Matches' range to a table..." -ForegroundColor Yellow
$listObj1 = $worksheet1.ListObjects.Add(1, $range1, $null, 1)
$listObj1.Name = "NationalNumberTable"
$listObj1.TableStyle = "TableStyleLight9"

# -------------------------
# CREATE "UNIQUE NUMBERS" SHEET
# -------------------------
Write-Host "Creating 'Unique Numbers' sheet..." -ForegroundColor Yellow
$worksheet2 = $workbook.Worksheets.Add()
$worksheet2.Name = "Unique Numbers"
$worksheet2.Columns.Item(1).NumberFormat = "@"

# Remove duplicates: group by NationalNumber and take the first occurrence.
$uniqueResults = $matchingResults | Group-Object NationalNumber | ForEach-Object { $_.Group[0] }
Write-Host "Found $($uniqueResults.Count) unique national number(s)." -ForegroundColor Green

$row2 = 1
$worksheet2.Cells.Item($row2, 1).Value2 = "National Number"
$worksheet2.Cells.Item($row2, 2).Value2 = "File Path"
$row2++

$i = 0
$totalUnique = $uniqueResults.Count
Write-Host "Writing data to 'Unique Numbers' sheet..." -ForegroundColor Yellow
foreach ($result in $uniqueResults) {
    $i++
    Write-Progress -Activity "Writing to Unique Numbers sheet" -Status "Row $i of $totalUnique" -PercentComplete (($i / $totalUnique) * 100)
    $worksheet2.Cells.Item($row2, 1).Value2 = $result.NationalNumber
    $worksheet2.Cells.Item($row2, 2).Value2 = $result.FilePath
    $row2++
}

$lastRow2 = $row2 - 1
$range2 = $worksheet2.Range("A1:B$lastRow2")
Write-Host "Converting 'Unique Numbers' range to a table..." -ForegroundColor Yellow
$listObj2 = $worksheet2.ListObjects.Add(1, $range2, $null, 1)
$listObj2.Name = "UniqueNationalNumberTable"
$listObj2.TableStyle = "TableStyleLight9"

# -------------------------
# REMOVE EXTRA SHEETS (if any)
# -------------------------
Write-Host "Cleaning up extra worksheets..." -ForegroundColor Yellow
# Create an array of allowed sheet names.
$allowedSheets = @("All Matches", "Unique Numbers")
# Iterate over a static copy of the worksheets collection.
foreach ($ws in @($workbook.Worksheets)) {
    if ($allowedSheets -notcontains $ws.Name) {
        Write-Host "Deleting extra sheet: $($ws.Name)" -ForegroundColor Magenta
        $ws.Delete()
    }
}

# -------------------------
# SAVE AND CLEAN UP
# -------------------------
$outputPath = Join-Path $rootDir "matching_files.xlsx"
Write-Host "Saving workbook to: $outputPath" -ForegroundColor Cyan
$workbook.SaveAs($outputPath)
$workbook.Close()
$excel.Quit()

Write-Host "=== Excel file created at: $outputPath ===" -ForegroundColor Cyan
Read-Host -Prompt "Press Enter to exit"
