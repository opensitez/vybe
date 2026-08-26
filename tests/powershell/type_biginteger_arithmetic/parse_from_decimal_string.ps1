# vybe-test: powershell/type_biginteger_arithmetic/parse_from_decimal_string
$str = "987654321098765432109876543210"
$val = [bigint]::Parse($str)
if ($val.ToString() -ne $str) {
    Write-Host "FAIL: expected $str, got $($val.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
