# vybe-test: powershell/math_trigonometric_functions/cos_zero_and_pi
$cos0 = [math]::Cos(0.0)
$cosPi = [math]::Cos([math]::PI)
if ($cos0 -ne 1.0 -or [math]::Abs($cosPi - (-1.0)) -gt 1e-12) {
    Write-Host "FAIL: Cos(0) and Cos(pi) failed"
    exit 1
}
Write-Host "PASS"
exit 0
