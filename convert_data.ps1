
$csvPath = "$(Get-Location)\data.csv"

# Attempt to detect encoding or just try UTF-8 first because we prefer Format 62
# But Format 62 usually adds BOM, so .NET UTF8 encoding should handle it.
# If fallback (Format 6) was used, it's Shift-JIS.

# Let's verify file preamble to guess encoding
$bytes = [System.IO.File]::ReadAllBytes($csvPath)
$isUtf8 = $false
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $isUtf8 = $true
}

$enc = $null
if ($isUtf8) {
    $enc = [System.Text.Encoding]::UTF8
    Write-Host "Detected UTF-8 BOM"
}
else {
    # If no BOM, assume Shift-JIS (or it might be UTF-8 no BOM, but Excel usually adds BOM)
    # Or checking if Format 62 was used but failed? 
    # Let's default to Shift-JIS if no BOM, as Format 6 produces that.
    $enc = [System.Text.Encoding]::GetEncoding(932)
    Write-Host "No BOM detected, assuming Shift-JIS"
}

$content = [System.IO.File]::ReadAllText($csvPath, $enc)

# Parse CSV
$rows = $content | ConvertFrom-Csv -Header c0, c1, c2, c3, c4, c5, c6 

$items = @()
$currentItem = $null

foreach ($row in $rows) {
    if ($row.c0 -match "^\d") {
        $currentItem = @{
            id              = $row.c0
            title           = $row.c1
            price_wholesale = $row.c3
            price_retail    = ""
        }
    }
    elseif ($currentItem -ne $null) {
        if ($row.c3 -ne $null -and $row.c3 -ne "") {
            $currentItem.price_retail = $row.c3
        }
        
        if ($currentItem.price_wholesale) {
            # Remove only currency symbols and commas, keep text if any? No, price is number.
            $currentItem.price_wholesale = $currentItem.price_wholesale -replace '[\\,¥"\s]', ''
        }
        if ($currentItem.price_retail) {
            $currentItem.price_retail = $currentItem.price_retail -replace '[\\,¥"\s]', ''
        }
        else {
            if (!$currentItem.price_retail) { $currentItem.price_retail = "0" }
        }
        
        try {
            $currentItem.price_wholesale = [int]$currentItem.price_wholesale
            $currentItem.price_retail = [int]$currentItem.price_retail
        }
        catch {}
        
        if ($currentItem.title -and $currentItem.title.Trim() -ne "") {
            $items += $currentItem
        }
        $currentItem = $null
    }
}

# Output to JS
$json = $items | ConvertTo-Json -Depth 3
$jsContent = "const DB_DATA = $json;"
[System.IO.File]::WriteAllText("$(Get-Location)\database.js", $jsContent, [System.Text.Encoding]::UTF8)

Write-Host "Converted $($items.Count) items to database.js"
