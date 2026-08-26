# vybe-test: powershell/math_logarithmic_and_exponential/sqrt_of_negative_returns_nan
$nan = [math]::Sqrt(-1.0)
if (-not [double]::IsNaN($nan)) {
    Write-Host "FAIL: Sqrt(-1) should return NaN"
    exit 1
}
Write-Host "PASS"
exit 0
