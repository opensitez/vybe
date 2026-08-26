# vybe-test: powershell/type_half_precision_floats/half_is_integer_check
$h = [System.Half]::Parse("8.0", [System.Globalization.CultureInfo]::InvariantCulture)
if (-not [System.Half]::IsInteger($h)) { Write-Host "FAIL: Half IsInteger failed"; exit 1 }
Write-Host "PASS"; exit 0
