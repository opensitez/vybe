# vybe-test: powershell/type_complex_numbers_arithmetic/complex_hashcode_consistency
$c1 = [System.Numerics.Complex]::new(7.0, 8.0)
$c2 = [System.Numerics.Complex]::new(7.0, 8.0)
if ($c1.GetHashCode() -ne $c2.GetHashCode()) { Write-Host "FAIL: Complex HashCode failed"; exit 1 }
Write-Host "PASS"; exit 0
