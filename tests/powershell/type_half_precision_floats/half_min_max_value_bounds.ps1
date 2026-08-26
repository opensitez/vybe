# vybe-test: powershell/type_half_precision_floats/half_min_max_value_bounds
$min = [System.Half]::MinValue
$max = [System.Half]::MaxValue
if ([double]$min -ge [double]$max -or [double]$min -gt -60000) { Write-Host "FAIL: Half bounds failed"; exit 1 }
Write-Host "PASS"; exit 0
