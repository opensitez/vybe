# vybe-test: powershell/math_logarithmic_and_exponential/pow_negative_exponent
$inv = [math]::Pow(2.0, -2.0)
if ($inv -ne 0.25) {
    Write-Host "FAIL: Pow negative exponent expected 0.25, got $inv"
    exit 1
}
Write-Host "PASS"
exit 0
