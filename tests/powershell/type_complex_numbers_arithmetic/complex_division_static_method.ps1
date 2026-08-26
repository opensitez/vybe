# vybe-test: powershell/type_complex_numbers_arithmetic/complex_division_static_method
$c1 = [System.Numerics.Complex]::new(-5.0, 10.0)
$c2 = [System.Numerics.Complex]::new(1.0, 2.0)
$div = [System.Numerics.Complex]::Divide($c1, $c2)
if ($div.Real -ne 3.0 -or $div.Imaginary -ne 4.0) { Write-Host "FAIL: Complex Divide failed"; exit 1 }
Write-Host "PASS"; exit 0
