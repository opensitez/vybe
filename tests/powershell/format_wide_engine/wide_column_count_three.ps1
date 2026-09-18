# vybe-test: powershell/format_wide_engine/wide_column_count_three
# Format-Wide -Column 3 formats output across 3 columns
$output = (1..9) | Format-Wide -Column 3 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# All elements 1 through 9 must appear
foreach ($n in 1..9) {
    if (-not ($output -match "\b$n\b")) {
        Write-Host "FAIL: number $n missing from wide 3-column output: $output"
        exit 1
    }
}

Write-Host "PASS"
exit 0
