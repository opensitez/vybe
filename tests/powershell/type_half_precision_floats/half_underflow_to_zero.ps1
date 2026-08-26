# vybe-test: powershell/type_half_precision_floats/half_underflow_to_zero
$tiny = 1e-10
$h = [System.Half]$tiny
if ([double]$h -ne 0) { Write-Host "FAIL: Half underflow to zero failed"; exit 1 }
Write-Host "PASS"; exit 0
