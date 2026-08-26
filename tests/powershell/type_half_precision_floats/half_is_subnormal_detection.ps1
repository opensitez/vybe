# vybe-test: powershell/type_half_precision_floats/half_is_subnormal_detection
$eps = [System.Half]::Epsilon
if (-not [System.Half]::IsSubnormal($eps)) { Write-Host "FAIL: Half IsSubnormal failed"; exit 1 }
Write-Host "PASS"; exit 0
