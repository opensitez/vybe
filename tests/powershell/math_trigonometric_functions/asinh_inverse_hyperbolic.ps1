# vybe-test: powershell/math_trigonometric_functions/asinh_inverse_hyperbolic
$x = 2.0
$asinh = [math]::Asinh($x)
$recovered = [math]::Sinh($asinh)
if ([math]::Abs($recovered - 2.0) -gt 1e-12) {
    Write-Host "FAIL: Asinh roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
