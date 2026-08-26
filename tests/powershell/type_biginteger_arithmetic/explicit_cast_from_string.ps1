# vybe-test: powershell/type_biginteger_arithmetic/explicit_cast_from_string
$str = "12345678901234567890"
$val = [bigint]$str
if ($val.GetType().Name -ne "BigInteger" -or $val.ToString() -ne $str) {
    Write-Host "FAIL: cast to BigInteger failed"
    exit 1
}
Write-Host "PASS"
exit 0
