# vybe-test: powershell/math_logarithmic_and_exponential/log_of_zero_returns_negative_infinity
$inf = [math]::Log(0.0)
if (-not [double]::IsNegativeInfinity($inf)) {
    Write-Host "FAIL: Log(0) should return NegativeInfinity"
    exit 1
}
Write-Host "PASS"
exit 0
