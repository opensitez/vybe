# vybe-test: powershell/math_trigonometric_functions/tanh_zero_and_limits
$tanh0 = [math]::Tanh(0.0)
$tanhLarge = [math]::Tanh(100.0)
if ($tanh0 -ne 0.0 -or [math]::Abs($tanhLarge - 1.0) -gt 1e-12) {
    Write-Host "FAIL: Tanh calculations failed"
    exit 1
}
Write-Host "PASS"
exit 0
