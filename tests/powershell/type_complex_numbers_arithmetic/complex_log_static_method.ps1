# vybe-test: powershell/type_complex_numbers_arithmetic/complex_log_static_method
$o = [System.Numerics.Complex]::One
$log = [System.Numerics.Complex]::Log($o)
if ($log.Real -ne 0.0 -or $log.Imaginary -ne 0.0) { Write-Host "FAIL: Complex Log failed"; exit 1 }
Write-Host "PASS"; exit 0
