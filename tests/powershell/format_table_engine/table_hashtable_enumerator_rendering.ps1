# vybe-test: powershell/format_table_engine/table_hashtable_enumerator_rendering
$hash = @{
    Region   = "East"
    Capacity = 500
}

# In PowerShell, piping hashtable GetEnumerator() produces formatted Name/Value table columns
$output = $hash.GetEnumerator() | Format-Table | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# Default format table view for DictionaryEntry renders Name and Value column headers
if (-not ($output -match "Name" -and $output -match "Value" -and $output -match "Region" -and $output -match "East")) {
    Write-Host "FAIL: hashtable enumerator failed to format into Name/Value table: $output"
    exit 1
}

Write-Host "PASS"
exit 0
