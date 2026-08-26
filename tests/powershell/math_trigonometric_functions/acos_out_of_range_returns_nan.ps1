# vybe-test: powershell/math_trigonometric_functions/acos_out_of_range_returns_nan
$nan = [math]::Acos(-1.5)
if (-not [double]::IsNaN($nan)) {
    Write-Host "FAIL: Acos(-1.5) should return NaN"
    exit 1
}
Write-Host "PASS"
exit 0
