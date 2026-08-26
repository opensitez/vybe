# vybe-test: powershell/math_ieee_special_values/single_precision_float_nan_and_infinity
$fNan = [float]::NaN
$fInf = [float]::PositiveInfinity
if (-not [float]::IsNaN($fNan) -or -not [float]::IsPositiveInfinity($fInf)) {
    Write-Host "FAIL: Single precision float special values failed"
    exit 1
}
Write-Host "PASS"
exit 0
