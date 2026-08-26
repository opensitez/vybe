# vybe-test: powershell/type_biginteger_arithmetic/bitwise_xor_toggle
$a = [bigint]255
$b = [bigint]15
$c = $a -bxor $b
if ($c -ne ([bigint]240)) {
    Write-Host "FAIL: expected 240, got $c"
    exit 1
}
Write-Host "PASS"
exit 0
