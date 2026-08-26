# vybe-test: powershell/math_logarithmic_and_exponential/exponential_and_log_inverse_relationship
$x = 4.567
$recovered = [math]::Log([math]::Exp($x))
if ([math]::Abs($recovered - $x) -gt 1e-12) {
    Write-Host "FAIL: Log(Exp(x)) inverse relationship failed"
    exit 1
}
Write-Host "PASS"
exit 0
