# vybe-test: powershell/math_trigonometric_functions/atanh_inverse_hyperbolic
$x = 0.5
$atanh = [math]::Atanh($x)
$recovered = [math]::Tanh($atanh)
if ([math]::Abs($recovered - 0.5) -gt 1e-12) {
    Write-Host "FAIL: Atanh roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
