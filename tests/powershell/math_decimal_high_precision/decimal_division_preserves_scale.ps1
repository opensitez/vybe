# vybe-test: powershell/math_decimal_high_precision/decimal_division_preserves_scale
[decimal]$a = 10
[decimal]$b = 4
$res = $a / $b
if ($res -ne [decimal]2.5) {
    Write-Host "FAIL: Decimal division expected 2.5, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
