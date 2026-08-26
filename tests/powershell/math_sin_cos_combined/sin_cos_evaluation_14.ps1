# vybe-test: powershell/math_sin_cos_combined/sin_cos_evaluation_14
$rad = [math]::PI / 4.0
$sin = [math]::Sin($rad)
$cos = [math]::Cos($rad)
if ([math]::Abs($sin - $cos) -gt 1e-6) { Write-Host "FAIL: Sin/Cos check failed"; exit 1 }
Write-Host "PASS"; exit 0
