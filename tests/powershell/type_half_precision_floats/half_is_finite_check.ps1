# vybe-test: powershell/type_half_precision_floats/half_is_finite_check
$h = [System.Half]::Parse("42", [System.Globalization.CultureInfo]::InvariantCulture)
if (-not [System.Half]::IsFinite($h) -or [System.Half]::IsInfinity($h)) { Write-Host "FAIL: Half IsFinite failed"; exit 1 }
Write-Host "PASS"; exit 0
