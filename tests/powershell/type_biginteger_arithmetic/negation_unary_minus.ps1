# vybe-test: powershell/type_biginteger_arithmetic/negation_unary_minus
$val = [bigint]123456789
$neg = -$val
if ($neg -ne [bigint]-123456789) {
    Write-Host "FAIL: expected -123456789, got $neg"
    exit 1
}
Write-Host "PASS"
exit 0
