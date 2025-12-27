
$files = Get-ChildItem ".\*.xlsx"
if ($files.Count -eq 0) {
    Write-Host "No Excel file found."
    exit
}
$file = $files[0]
$path = $file.FullName

Write-Host "Opening Excel: $path"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1)
    
    # Locate last row
    $lastRow = $sheet.UsedRange.Rows.Count
    Write-Host "Reading $lastRow rows..."

    # Read all values into a 2D array (much faster than cell by cell)
    $range = $sheet.Range("A1:G$lastRow")
    $values = $range.Value2

    # $values is a 1-based 2D array in PowerShell [row, col]
    # ComObject array indexing is 1-based usually
    
    $items = @()
    $currentItem = $null
    
    # Iterate safely
    for ($r = 1; $r -le $lastRow; $r++) {
        $c0 = $values[$r, 1] # ID column
        $c1 = $values[$r, 2] # Title
        $c3 = $values[$r, 4] # Price
        
        # Check if ID exists (Main row)
        if ($c0 -match "^\d") {
            $currentItem = @{
                id              = $c0
                title           = $c1
                price_wholesale = $c3
                price_retail    = ""
            }
        }
        elseif ($currentItem -ne $null) {
            # Retail price row (next row)
            if ($c3) {
                $currentItem.price_retail = $c3
            }
            
            # Add to list
            if ($currentItem.title) {
                # Clean up prices (if they are strings)
                if ($currentItem.price_wholesale -is [string]) {
                    $currentItem.price_wholesale = $currentItem.price_wholesale -replace '[\\,¥"\s]', ''
                }
                if ($currentItem.price_retail -is [string]) {
                    $currentItem.price_retail = $currentItem.price_retail -replace '[\\,¥"\s]', ''
                }
                
                # Convert to Int
                try { $currentItem.price_wholesale = [int]$currentItem.price_wholesale } catch {}
                try { $currentItem.price_retail = [int]$currentItem.price_retail } catch {}
                
                $items += $currentItem
            }
            $currentItem = $null
        }
    }
    
    $workbook.Close($false)
    
    # Output to JS
    $json = $items | ConvertTo-Json -Depth 3
    $jsContent = "const DB_DATA = $json;"
    [System.IO.File]::WriteAllText("$(Get-Location)\database.js", $jsContent, [System.Text.Encoding]::UTF8)

    Write-Host "Success: Generated database.js with $($items.Count) items."

}
catch {
    Write-Host "Error: $_"
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
