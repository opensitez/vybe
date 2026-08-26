# vybe-test: powershell/math_decimal_high_precision/decimal_truncate_method
[decimal]$d = 15.987
$t = [decimal]::Truncate($d)
if ($t -ne [decimal]15) {
    Write-Host "FAIL: Decimal Truncate failed, got $t"
    exit 1
}
Write-Host "PASS"
exit 0
