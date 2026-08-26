# vybe-test: powershell/type_half_precision_floats/half_equality_same_values
$h1 = [System.Half]::Parse("5.5", [System.Globalization.CultureInfo]::InvariantCulture)
$h2 = [System.Half]::Parse("5.5", [System.Globalization.CultureInfo]::InvariantCulture)
if (-not $h1.Equals($h2)) { Write-Host "FAIL: Half Equals failed"; exit 1 }
Write-Host "PASS"; exit 0
