# vybe-test: powershell/type_half_precision_floats/half_compareto_ordering
$h1 = [System.Half]::Parse("10.0", [System.Globalization.CultureInfo]::InvariantCulture)
$h2 = [System.Half]::Parse("20.0", [System.Globalization.CultureInfo]::InvariantCulture)
if ($h1.CompareTo($h2) -ge 0 -or $h2.CompareTo($h1) -le 0) { Write-Host "FAIL: Half CompareTo failed"; exit 1 }
Write-Host "PASS"; exit 0
