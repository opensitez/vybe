# vybe-test: powershell/type_biginteger_arithmetic/bitwise_or_combine
$a = [bigint]16
$b = [bigint]8
$c = $a -bor $b
if ($c -ne ([bigint]24)) {
    Write-Host "FAIL: expected 24, got $c"
    exit 1
}
Write-Host "PASS"
exit 0
