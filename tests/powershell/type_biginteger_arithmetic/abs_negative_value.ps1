# vybe-test: powershell/type_biginteger_arithmetic/abs_negative_value
$val = [bigint]::Parse("-999999999999999999999")
$abs = [bigint]::Abs($val)
$expected = [bigint]::Parse("999999999999999999999")
if ($abs -ne $expected) {
    Write-Host "FAIL: expected $expected, got $abs"
    exit 1
}
Write-Host "PASS"
exit 0
