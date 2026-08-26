# vybe-test: powershell/math_ieee_special_values/parse_nan_string
$val = [double]::Parse("NaN")
if (-not [double]::IsNaN($val)) {
    Write-Host "FAIL: Parse NaN string failed"
    exit 1
}
Write-Host "PASS"
exit 0
