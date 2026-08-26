# vybe-test: powershell/math_trigonometric_functions/acosh_inverse_hyperbolic
$x = 3.0
$acosh = [math]::Acosh($x)
$recovered = [math]::Cosh($acosh)
if ([math]::Abs($recovered - 3.0) -gt 1e-12) {
    Write-Host "FAIL: Acosh roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
