# vybe-test: powershell/type_biginteger_arithmetic/greatest_common_divisor
$a = [bigint]54
$b = [bigint]24
$gcd = [bigint]::GreatestCommonDivisor($a, $b)
if ($gcd -ne [bigint]6) {
    Write-Host "FAIL: GCD(54, 24) expected 6, got $gcd"
    exit 1
}
Write-Host "PASS"
exit 0
