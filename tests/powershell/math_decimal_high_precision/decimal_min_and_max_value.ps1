# vybe-test: powershell/math_decimal_high_precision/decimal_min_and_max_value
$min = [decimal]::MinValue
$max = [decimal]::MaxValue
if ($min -ge 0 -or $max -le 0) {
    Write-Host "FAIL: Decimal Min/Max value failed"
    exit 1
}
Write-Host "PASS"
exit 0
