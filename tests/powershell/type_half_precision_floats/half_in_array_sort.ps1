# vybe-test: powershell/type_half_precision_floats/half_in_array_sort
$h1 = [System.Half]::Parse("30", [System.Globalization.CultureInfo]::InvariantCulture)
$h2 = [System.Half]::Parse("10", [System.Globalization.CultureInfo]::InvariantCulture)
$h3 = [System.Half]::Parse("20", [System.Globalization.CultureInfo]::InvariantCulture)
$arr = [System.Half[]]@($h1, $h2, $h3)
[System.Array]::Sort($arr)
if ([double]$arr[0] -ne 10 -or [double]$arr[2] -ne 30) { Write-Host "FAIL: Half Array.Sort failed"; exit 1 }
Write-Host "PASS"; exit 0
