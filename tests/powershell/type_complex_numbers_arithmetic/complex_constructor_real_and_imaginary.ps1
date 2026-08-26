# vybe-test: powershell/type_complex_numbers_arithmetic/complex_constructor_real_and_imaginary
$c = [System.Numerics.Complex]::new(3.0, 4.0)
if ($c.Real -ne 3.0 -or $c.Imaginary -ne 4.0) { Write-Host "FAIL: Complex constructor failed"; exit 1 }
Write-Host "PASS"; exit 0
