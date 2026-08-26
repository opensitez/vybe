# vybe-test: powershell/math_decimal_high_precision/decimal_modulo_operation
[decimal]$a = 10.5
[decimal]$b = 3.0
$rem = $a % $b
if ($rem -ne [decimal]1.5) {
    Write-Host "FAIL: Decimal modulo expected 1.5, got $rem"
    exit 1
}
Write-Host "PASS"
exit 0
