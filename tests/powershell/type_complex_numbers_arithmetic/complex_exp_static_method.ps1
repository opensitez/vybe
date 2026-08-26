# vybe-test: powershell/type_complex_numbers_arithmetic/complex_exp_static_method
$z = [System.Numerics.Complex]::Zero
$exp = [System.Numerics.Complex]::Exp($z)
if ($exp.Real -ne 1.0 -or $exp.Imaginary -ne 0.0) { Write-Host "FAIL: Complex Exp failed"; exit 1 }
Write-Host "PASS"; exit 0
