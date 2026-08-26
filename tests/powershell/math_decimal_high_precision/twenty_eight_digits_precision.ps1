# vybe-test: powershell/math_decimal_high_precision/twenty_eight_digits_precision
$str = "1234567890123456789012345678"
[decimal]$d = [decimal]::Parse($str)
if ($d.ToString() -ne $str) {
    Write-Host "FAIL: 28 digits precision parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
