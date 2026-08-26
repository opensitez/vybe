# vybe-test: powershell/type_half_precision_floats/half_conversion_from_double
$d = 100.25
$h = [System.Half]$d
if ([double]$h -ne 100.25) { Write-Host "FAIL: Half double conversion failed"; exit 1 }
Write-Host "PASS"; exit 0
