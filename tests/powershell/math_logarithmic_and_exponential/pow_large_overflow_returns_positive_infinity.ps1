# vybe-test: powershell/math_logarithmic_and_exponential/pow_large_overflow_returns_positive_infinity
$inf = [math]::Pow(10.0, 400.0)
if (-not [double]::IsPositiveInfinity($inf)) {
    Write-Host "FAIL: 10^400 expected PositiveInfinity"
    exit 1
}
Write-Host "PASS"
exit 0
