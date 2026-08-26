# vybe-test: powershell/type_half_precision_floats/half_parse_from_string
$h = [System.Half]::Parse("3.14", [System.Globalization.CultureInfo]::InvariantCulture)
$d = [double]$h
if ($d -lt 3.13 -or $d -gt 3.15) { Write-Host "FAIL: Half Parse failed"; exit 1 }
Write-Host "PASS"; exit 0
