# vybe-test: powershell/format_wide_engine/wide_primitive_positive_integers
# Primitive numbers format cleanly in wide columns
$output = (100..105) | Format-Wide -Column 3 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

foreach ($n in 100..105) {
    if (-not ($output -match "\b$n\b")) {
        Write-Host "FAIL: integer $n missing from wide output: $output"
        exit 1
    }
}

Write-Host "PASS"
exit 0
