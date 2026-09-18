# vybe-test: powershell/format_wide_engine/wide_handles_empty_collection_input
$emptyArr = @()

# Piping empty collection into Format-Wide must succeed and emit empty string
$output = $emptyArr | Format-Wide | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if ($output.Trim().Length -ne 0) {
    Write-Host "FAIL: expected empty string for empty input, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
