# vybe-test: powershell/type_complex_numbers_arithmetic/complex_sqrt_static_method
$c = [System.Numerics.Complex]::new(-4.0, 0.0)
$sqrt = [System.Numerics.Complex]::Sqrt($c)
if ($sqrt.Real -ne 0.0 -or $sqrt.Imaginary -ne 2.0) { Write-Host "FAIL: Complex Sqrt failed, got real=$($sqrt.Real), imag=$($sqrt.Imaginary)"; exit 1 }
Write-Host "PASS"; exit 0
