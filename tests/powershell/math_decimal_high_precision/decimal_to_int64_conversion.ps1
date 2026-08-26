# vybe-test: powershell/math_decimal_high_precision/decimal_to_int64_conversion
[decimal]$d = 1234567890123
$i = [int64]$d
if ($i -ne 1234567890123) {
    Write-Host "FAIL: Decimal to int64 cast failed"
    exit 1
}
Write-Host "PASS"
exit 0
