# vybe-test: powershell/type_half_precision_floats/half_conversion_from_single
$s = [float]12.5
$h = [System.Half]$s
$back = [float]$h
if ($back -ne 12.5) { Write-Host "FAIL: Half float conversion failed"; exit 1 }
Write-Host "PASS"; exit 0
