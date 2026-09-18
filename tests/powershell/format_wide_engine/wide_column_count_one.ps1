# vybe-test: powershell/format_wide_engine/wide_column_count_one
# Format-Wide -Column 1 formats output into a single vertical column
$output = (1..4) | Format-Wide -Column 1 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

$nonEmptyLines = $output.Trim().Split("`n") | Where-Object { $_.Trim().Length -gt 0 }

# Exactly 4 lines of output for 4 single-column items
if ($nonEmptyLines.Count -ne 4) {
    Write-Host "FAIL: expected exactly 4 lines for 4 items in 1-column wide view, got: $($nonEmptyLines.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
