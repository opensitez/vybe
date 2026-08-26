# vybe-test: powershell/type_complex_numbers_arithmetic/complex_conjugate_method
$c = [System.Numerics.Complex]::new(5.0, -2.0)
$conj = [System.Numerics.Complex]::Conjugate($c)
if ($conj.Real -ne 5.0 -or $conj.Imaginary -ne 2.0) { Write-Host "FAIL: Complex Conjugate failed"; exit 1 }
Write-Host "PASS"; exit 0
