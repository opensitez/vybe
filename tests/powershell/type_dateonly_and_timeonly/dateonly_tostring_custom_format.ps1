# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_tostring_custom_format
$d = [System.DateOnly]::new(2026, 8, 26)
$str = $d.ToString("yyyy/MM/dd", [System.Globalization.CultureInfo]::InvariantCulture)
if ($str -ne "2026/08/26") { Write-Host "FAIL: DateOnly custom format failed, got $str"; exit 1 }
Write-Host "PASS"; exit 0
