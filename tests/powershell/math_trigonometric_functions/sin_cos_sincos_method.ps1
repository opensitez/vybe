# vybe-test: powershell/math_trigonometric_functions/sin_cos_sincos_method
$theta = 0.5
$sin = [math]::Sin($theta)
$cos = [math]::Cos($theta)
if ([math]::Abs($sin - 0.4794255386) -gt 1e-8 -or [math]::Abs($cos - 0.8775825618) -gt 1e-8) {
    Write-Host "FAIL: Sin/Cos calculation values failed"
    exit 1
}
Write-Host "PASS"
exit 0
