# vybe-test: powershell/math_ieee_special_values/isnan_check_true_and_false
$nan = [double]::NaN
$num = 42.0
if (-not [double]::IsNaN($nan) -or [double]::IsNaN($num)) {
    Write-Host "FAIL: IsNaN check failed"
    exit 1
}
Write-Host "PASS"
exit 0
