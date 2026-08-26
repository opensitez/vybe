# vybe-test: powershell/math_logarithmic_and_exponential/log_of_negative_number_returns_nan
$nan = [math]::Log(-5.0)
if (-not [double]::IsNaN($nan)) {
    Write-Host "FAIL: Log(-5) should return NaN"
    exit 1
}
Write-Host "PASS"
exit 0
