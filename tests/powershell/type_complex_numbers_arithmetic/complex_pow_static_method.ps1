# vybe-test: powershell/type_complex_numbers_arithmetic/complex_pow_static_method
$i = [System.Numerics.Complex]::ImaginaryOne
$iSquared = [System.Numerics.Complex]::Pow($i, 2.0)
if ($iSquared.Real -ne -1.0 -or [math]::Abs($iSquared.Imaginary) -gt 1e-9) { Write-Host "FAIL: Complex Pow i^2 failed"; exit 1 }
Write-Host "PASS"; exit 0
