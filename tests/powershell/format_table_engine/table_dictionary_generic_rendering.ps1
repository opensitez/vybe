# vybe-test: powershell/format_table_engine/table_dictionary_generic_rendering
# Generic Dictionary[TKey, TValue] formatted into a table
$dict = [System.Collections.Generic.Dictionary[string, int]]::new()
$dict["Apples"] = 12
$dict["Oranges"] = 24

$output = $dict | Format-Table | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The table must contain Key and Value headers and the items
if (-not ($output -match "Key" -and $output -match "Value" -and $output -match "Apples" -and $output -match "12")) {
    Write-Host "FAIL: generic dictionary table rendering failed: $output"
    exit 1
}

Write-Host "PASS"
exit 0
