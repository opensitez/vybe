# vybe-test: powershell/type_half_precision_floats/half_tostring_formatting
$h = [System.Half]::Parse("7.5", [System.Globalization.CultureInfo]::InvariantCulture)
$str = $h.ToString([System.Globalization.CultureInfo]::InvariantCulture)
if ($str -ne "7.5") { Write-Host "FAIL: Half ToString failed, got $str"; exit 1 }
Write-Host "PASS"; exit 0
