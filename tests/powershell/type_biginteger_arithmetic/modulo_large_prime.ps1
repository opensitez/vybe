# vybe-test: powershell/type_biginteger_arithmetic/modulo_large_prime
$num = [bigint]::Parse("1000000000000000000000000000057")
$mod = [bigint]1000000007
$rem = $num % $mod
$expected = [bigint]999657064
if ($rem -ne $expected) {
    Write-Host "FAIL: expected $expected, got $rem"
    exit 1
}
Write-Host "PASS"
exit 0
