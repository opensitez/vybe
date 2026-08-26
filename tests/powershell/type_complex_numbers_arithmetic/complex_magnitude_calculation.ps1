# vybe-test: powershell/type_complex_numbers_arithmetic/complex_magnitude_calculation
$c = [System.Numerics.Complex]::new(3.0, 4.0)
if ($c.Magnitude -ne 5.0) { Write-Host "FAIL: Complex Magnitude expected 5.0, got $($c.Magnitude)"; exit 1 }
Write-Host "PASS"; exit 0
