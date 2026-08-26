# vybe-test: powershell/type_half_precision_floats/half_hashcode_consistency
$h1 = [System.Half]::Parse("99", [System.Globalization.CultureInfo]::InvariantCulture)
$h2 = [System.Half]::Parse("99", [System.Globalization.CultureInfo]::InvariantCulture)
if ($h1.GetHashCode() -ne $h2.GetHashCode()) { Write-Host "FAIL: Half GetHashCode failed"; exit 1 }
Write-Host "PASS"; exit 0
