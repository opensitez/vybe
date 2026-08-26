# vybe-test: powershell/math_ieee_special_values/nan_not_equal_to_itself
$nan = [double]::NaN
if ($nan -eq $nan) {
    Write-Host "FAIL: IEEE 754 NaN must not compare equal to itself"
    exit 1
}
Write-Host "PASS"
exit 0
