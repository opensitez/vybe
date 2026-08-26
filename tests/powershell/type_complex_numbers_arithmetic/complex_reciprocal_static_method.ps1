# vybe-test: powershell/type_complex_numbers_arithmetic/complex_reciprocal_static_method
$c = [System.Numerics.Complex]::new(2.0, 0.0)
$rec = [System.Numerics.Complex]::Reciprocal($c)
if ($rec.Real -ne 0.5 -or $rec.Imaginary -ne 0.0) { Write-Host "FAIL: Complex Reciprocal failed"; exit 1 }
Write-Host "PASS"; exit 0
