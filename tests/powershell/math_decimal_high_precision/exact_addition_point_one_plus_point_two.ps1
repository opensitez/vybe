# vybe-test: powershell/math_decimal_high_precision/exact_addition_point_one_plus_point_two
[decimal]$a = 0.1
[decimal]$b = 0.2
[decimal]$c = $a + $b
if ($c -ne [decimal]0.3) {
    Write-Host "FAIL: Decimal 0.1 + 0.2 expected exact 0.3, got $c"
    exit 1
}
Write-Host "PASS"
exit 0
