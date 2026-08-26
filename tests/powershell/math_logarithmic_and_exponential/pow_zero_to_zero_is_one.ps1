# vybe-test: powershell/math_logarithmic_and_exponential/pow_zero_to_zero_is_one
$p = [math]::Pow(0.0, 0.0)
if ($p -ne 1.0) {
    Write-Host "FAIL: 0^0 expected 1.0, got $p"
    exit 1
}
Write-Host "PASS"
exit 0
