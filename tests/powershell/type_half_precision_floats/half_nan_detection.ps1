# vybe-test: powershell/type_half_precision_floats/half_nan_detection
$nan = [System.Half]::NaN
if (-not [System.Half]::IsNaN($nan) -or [System.Half]::IsFinite($nan)) { Write-Host "FAIL: Half NaN failed"; exit 1 }
Write-Host "PASS"; exit 0
