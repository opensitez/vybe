# vybe-test: powershell/type_biginteger_arithmetic/subtraction_into_negative
$a = [bigint]::Parse("100000000000000000000000")
$b = [bigint]::Parse("300000000000000000000000")
$c = $a - $b
$expected = [bigint]::Parse("-200000000000000000000000")
if ($c -ne $expected) {
    Write-Host "FAIL: expected $expected, got $c"
    exit 1
}
Write-Host "PASS"
exit 0
