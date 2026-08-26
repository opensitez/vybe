# vybe-test: powershell/type_biginteger_arithmetic/division_exact_large
$num = [bigint]::Parse("123456789012345678901234567890")
$den = [bigint]10
$res = [bigint]::Divide($num, $den)
$expected = [bigint]::Parse("12345678901234567890123456789")
if ($res -ne $expected) {
    Write-Host "FAIL: expected $expected, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
