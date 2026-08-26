# vybe-test: powershell/type_complex_numbers_arithmetic/complex_addition_static_method
$c1 = [System.Numerics.Complex]::new(1.0, 2.0)
$c2 = [System.Numerics.Complex]::new(3.0, 4.0)
$sum = [System.Numerics.Complex]::Add($c1, $c2)
if ($sum.Real -ne 4.0 -or $sum.Imaginary -ne 6.0) { Write-Host "FAIL: Complex Add failed"; exit 1 }
Write-Host "PASS"; exit 0
