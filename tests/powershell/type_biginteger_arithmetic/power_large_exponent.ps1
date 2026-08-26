# vybe-test: powershell/type_biginteger_arithmetic/power_large_exponent
$base = [bigint]2
$pow = [bigint]::Pow($base, 100)
$expected = [bigint]::Parse("1267650600228229401496703205376")
if ($pow -ne $expected) {
    Write-Host "FAIL: 2^100 expected $expected, got $pow"
    exit 1
}
Write-Host "PASS"
exit 0
