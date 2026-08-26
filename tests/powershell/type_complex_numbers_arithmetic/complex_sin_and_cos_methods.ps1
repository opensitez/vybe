# vybe-test: powershell/type_complex_numbers_arithmetic/complex_sin_and_cos_methods
$z = [System.Numerics.Complex]::Zero
$sin = [System.Numerics.Complex]::Sin($z)
$cos = [System.Numerics.Complex]::Cos($z)
if ($sin.Real -ne 0.0 -or $cos.Real -ne 1.0) { Write-Host "FAIL: Complex Sin/Cos failed"; exit 1 }
Write-Host "PASS"; exit 0
