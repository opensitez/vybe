# vybe-test: powershell/math_trigonometric_functions/asin_out_of_range_returns_nan
$nan = [math]::Asin(2.0)
if (-not [double]::IsNaN($nan)) {
    Write-Host "FAIL: Asin(2.0) should return NaN"
    exit 1
}
Write-Host "PASS"
exit 0
