
$files = Get-ChildItem ".\*.xlsx"
if ($files.Count -eq 0) {
    Write-Host "No Excel file found."
    exit
}
$file = $files[0]
$path = $file.FullName
$csvPath = Join-Path (Get-Location) "data.csv"

Write-Host "Opening: $path"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($path)
    # Format 62 = CSV UTF-8 (Available in Excel 2016+)
    $workbook.SaveAs($csvPath, 62)
    $workbook.Close($false)
    Write-Host "Success: Saved as UTF-8 CSV"
}
catch {
    Write-Host "Error: $_"
    # Fallback to standard CSV (6) if 62 fails
    try {
        $workbook.SaveAs($csvPath, 6)
        Write-Host "Fallback: Saved as Standard CSV (Shift-JIS)"
    }
    catch {
        Write-Host "Error in fallback: $_"
    }
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
