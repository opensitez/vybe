# vybe-test: powershell/type_half_precision_floats/half_epsilon_property
$eps = [System.Half]::Epsilon
if ([double]$eps -le 0) { Write-Host "FAIL: Half Epsilon failed"; exit 1 }
Write-Host "PASS"; exit 0
