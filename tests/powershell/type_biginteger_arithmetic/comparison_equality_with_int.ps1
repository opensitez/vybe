# vybe-test: powershell/type_biginteger_arithmetic/comparison_equality_with_int
$a = [bigint]42
$b = 42
if ($a -ne $b) {
    Write-Host "FAIL: bigint 42 should equal int 42"
    exit 1
}
Write-Host "PASS"
exit 0
