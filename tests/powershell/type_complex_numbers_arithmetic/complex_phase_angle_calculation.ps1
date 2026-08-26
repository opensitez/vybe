# vybe-test: powershell/type_complex_numbers_arithmetic/complex_phase_angle_calculation
$c = [System.Numerics.Complex]::new(0.0, 1.0)
$halfPi = [math]::PI / 2.0
if ([math]::Abs($c.Phase - $halfPi) -gt 1e-9) { Write-Host "FAIL: Complex Phase expected pi/2, got $($c.Phase)"; exit 1 }
Write-Host "PASS"; exit 0
