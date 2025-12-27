
$csvPath = "$(Get-Location)\data.csv"

# Try reading as UTF-8 first
try {
    # Check for BOM or just try UTF8
    $enc = [System.Text.Encoding]::UTF8
    $content = [System.IO.File]::ReadAllText($csvPath, $enc)
    
    # Simple check: if we see lots of replacement chars or garbage, maybe it wasn't UTF8
    # But let's just output head to console for debug
    Write-Host "Debug Head (UTF8): $( $content.Substring(0, [math]::Min(50, $content.Length)) )"
}
catch {
    Write-Host "Read error"
}
