# vybe-test: powershell/type_complex_numbers_arithmetic/complex_frompolarcoordinates_factory
$polar = [System.Numerics.Complex]::FromPolarCoordinates(5.0, 0.0)
if ($polar.Real -ne 5.0 -or $polar.Imaginary -ne 0.0) { Write-Host "FAIL: FromPolarCoordinates failed"; exit 1 }
Write-Host "PASS"; exit 0
