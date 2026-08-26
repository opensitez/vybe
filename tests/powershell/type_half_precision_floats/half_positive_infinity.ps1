# vybe-test: powershell/type_half_precision_floats/half_positive_infinity
$inf = [System.Half]::PositiveInfinity
if (-not [System.Half]::IsPositiveInfinity($inf) -or [System.Half]::IsNegativeInfinity($inf)) { Write-Host "FAIL: Half PositiveInfinity failed"; exit 1 }
Write-Host "PASS"; exit 0
