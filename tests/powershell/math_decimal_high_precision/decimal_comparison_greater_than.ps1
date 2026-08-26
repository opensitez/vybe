# vybe-test: powershell/math_decimal_high_precision/decimal_comparison_greater_than
[decimal]$d1 = 100.001
[decimal]$d2 = 100.0009
if (-not ($d1 -gt $d2)) {
    Write-Host "FAIL: Decimal precision comparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
