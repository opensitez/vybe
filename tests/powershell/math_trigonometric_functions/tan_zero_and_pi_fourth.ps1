# vybe-test: powershell/math_trigonometric_functions/tan_zero_and_pi_fourth
$tan0 = [math]::Tan(0.0)
$tanPi4 = [math]::Tan([math]::PI / 4.0)
if ($tan0 -ne 0.0 -or [math]::Abs($tanPi4 - 1.0) -gt 1e-12) {
    Write-Host "FAIL: Tan(0) and Tan(pi/4) failed"
    exit 1
}
Write-Host "PASS"
exit 0
