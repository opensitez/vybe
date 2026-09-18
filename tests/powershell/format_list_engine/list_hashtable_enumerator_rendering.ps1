# vybe-test: powershell/format_list_engine/list_hashtable_enumerator_rendering
$hash = @{ DatabaseHost = "db.local" }

# Hashtable enumerator formatted in list view renders Name and Value pairs
$output = $hash.GetEnumerator() | Format-List | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "Name\s*:\s*DatabaseHost" -and $output -match "Value\s*:\s*db\.local")) {
    Write-Host "FAIL: hashtable enumerator failed to format in list view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
