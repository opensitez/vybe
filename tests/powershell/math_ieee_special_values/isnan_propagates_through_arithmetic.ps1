# vybe-test: powershell/math_ieee_special_values/isnan_propagates_through_arithmetic
$nan = [double]::NaN
$res = $nan + 10.0
if (-not [double]::IsNaN($res)) {
    Write-Host "FAIL: NaN arithmetic propagation failed"
    exit 1
}
Write-Host "PASS"
exit 0
