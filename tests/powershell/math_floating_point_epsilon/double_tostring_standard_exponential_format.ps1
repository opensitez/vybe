# vybe-test: powershell/math_floating_point_epsilon/double_tostring_standard_exponential_format
$val = 12345.0
$str = $val.ToString("E2", [System.Globalization.CultureInfo]::InvariantCulture)
if ($str -ne "1.23E+004" -and $str -ne "1.23E+04") {
    Write-Host "FAIL: Exponential format E2 failed, got $str"
    exit 1
}
Write-Host "PASS"
exit 0
