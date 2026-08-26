# vybe-test: powershell/type_complex_numbers_arithmetic/complex_subtraction_static_method
$c1 = [System.Numerics.Complex]::new(5.0, 7.0)
$c2 = [System.Numerics.Complex]::new(2.0, 3.0)
$diff = [System.Numerics.Complex]::Subtract($c1, $c2)
if ($diff.Real -ne 3.0 -or $diff.Imaginary -ne 4.0) { Write-Host "FAIL: Complex Subtract failed"; exit 1 }
Write-Host "PASS"; exit 0
