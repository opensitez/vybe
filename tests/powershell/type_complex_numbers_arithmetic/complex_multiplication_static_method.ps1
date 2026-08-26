# vybe-test: powershell/type_complex_numbers_arithmetic/complex_multiplication_static_method
$c1 = [System.Numerics.Complex]::new(1.0, 2.0)
$c2 = [System.Numerics.Complex]::new(3.0, 4.0)
$prod = [System.Numerics.Complex]::Multiply($c1, $c2) # (1*3 - 2*4) + (1*4 + 2*3)i = -5 + 10i
if ($prod.Real -ne -5.0 -or $prod.Imaginary -ne 10.0) { Write-Host "FAIL: Complex Multiply failed"; exit 1 }
Write-Host "PASS"; exit 0
