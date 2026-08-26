# vybe-test: powershell/type_complex_numbers_arithmetic/complex_equality_same_components
$c1 = [System.Numerics.Complex]::new(2.5, 3.5)
$c2 = [System.Numerics.Complex]::new(2.5, 3.5)
if (-not $c1.Equals($c2)) { Write-Host "FAIL: Complex Equals failed"; exit 1 }
Write-Host "PASS"; exit 0
