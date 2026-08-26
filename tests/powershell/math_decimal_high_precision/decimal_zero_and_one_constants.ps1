# vybe-test: powershell/math_decimal_high_precision/decimal_zero_and_one_constants
$z = [decimal]::Zero
$o = [decimal]::One
$m = [decimal]::MinusOne
if ($z -ne 0 -or $o -ne 1 -or $m -ne -1) {
    Write-Host "FAIL: Decimal constants failed"
    exit 1
}
Write-Host "PASS"
exit 0
