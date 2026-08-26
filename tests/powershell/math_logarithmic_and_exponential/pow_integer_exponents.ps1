# vybe-test: powershell/math_logarithmic_and_exponential/pow_integer_exponents
$p1 = [math]::Pow(2.0, 10.0)
$p2 = [math]::Pow(3.0, 3.0)
if ($p1 -ne 1024.0 -or $p2 -ne 27.0) {
    Write-Host "FAIL: Pow integer exponents failed"
    exit 1
}
Write-Host "PASS"
exit 0
