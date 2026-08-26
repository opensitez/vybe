# vybe-test: powershell/math_logarithmic_and_exponential/pow_fractional_exponent_square_root
$sqrt = [math]::Pow(16.0, 0.5)
if ($sqrt -ne 4.0) {
    Write-Host "FAIL: Pow fractional exponent expected 4.0, got $sqrt"
    exit 1
}
Write-Host "PASS"
exit 0
