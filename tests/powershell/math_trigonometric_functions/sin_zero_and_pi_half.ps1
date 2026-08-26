# vybe-test: powershell/math_trigonometric_functions/sin_zero_and_pi_half
$sin0 = [math]::Sin(0.0)
$sinHalfPi = [math]::Sin([math]::PI / 2.0)
if ($sin0 -ne 0.0 -or [math]::Abs($sinHalfPi - 1.0) -gt 1e-12) {
    Write-Host "FAIL: Sin(0) and Sin(pi/2) failed"
    exit 1
}
Write-Host "PASS"
exit 0
