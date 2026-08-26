# vybe-test: powershell/type_half_precision_floats/half_in_generic_list
$list = [System.Collections.Generic.List[System.Half]]::new()
$list.Add([System.Half]::Parse("1.1", [System.Globalization.CultureInfo]::InvariantCulture))
if ($list.Count -ne 1) { Write-Host "FAIL: Half in List failed"; exit 1 }
Write-Host "PASS"; exit 0
