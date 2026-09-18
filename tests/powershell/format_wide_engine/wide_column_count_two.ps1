# vybe-test: powershell/format_wide_engine/wide_column_count_two
# Format-Wide -Column 2 specifies formatting across exactly 2 columns
$output = (1..6) | Format-Wide -Column 2 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# All numbers 1 through 6 must be represented
foreach ($n in 1..6) {
    if (-not ($output -match "\b$n\b")) {
        Write-Host "FAIL: number $n missing from wide 2-column output: $output"
        exit 1
    }
}

Write-Host "PASS"
exit 0
