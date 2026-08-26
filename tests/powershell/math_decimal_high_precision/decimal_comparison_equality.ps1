# vybe-test: powershell/math_decimal_high_precision/decimal_comparison_equality
[decimal]$d1 = [decimal]::Parse("5.00")
[decimal]$d2 = [decimal]::Parse("5.0")
if ($d1 -ne $d2) {
    Write-Host "FAIL: Decimal trailing zeros should compare equal"
    exit 1
}
Write-Host "PASS"
exit 0
