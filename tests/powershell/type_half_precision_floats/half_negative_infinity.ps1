# vybe-test: powershell/type_half_precision_floats/half_negative_infinity
$ninf = [System.Half]::NegativeInfinity
if (-not [System.Half]::IsNegativeInfinity($ninf) -or [System.Half]::IsPositiveInfinity($ninf)) { Write-Host "FAIL: Half NegativeInfinity failed"; exit 1 }
Write-Host "PASS"; exit 0
