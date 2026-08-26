# vybe-test: powershell/type_complex_numbers_arithmetic/complex_zero_one_imaginaryone_constants
$z = [System.Numerics.Complex]::Zero
$o = [System.Numerics.Complex]::One
$i = [System.Numerics.Complex]::ImaginaryOne
if ($z.Real -ne 0 -or $o.Real -ne 1 -or $i.Imaginary -ne 1) { Write-Host "FAIL: Complex constants failed"; exit 1 }
Write-Host "PASS"; exit 0
