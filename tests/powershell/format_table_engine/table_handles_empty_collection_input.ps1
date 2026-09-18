# vybe-test: powershell/format_table_engine/table_handles_empty_collection_input
$emptyCollection = @()

# Piping an empty collection into Format-Table must succeed without errors and yield empty string
$output = $emptyCollection | Format-Table -Property PropA, PropB | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if ($output.Trim().Length -ne 0) {
    Write-Host "FAIL: expected empty output for empty input collection, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
