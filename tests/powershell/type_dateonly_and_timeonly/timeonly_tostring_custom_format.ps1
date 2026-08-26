# vybe-test: powershell/type_dateonly_and_timeonly/timeonly_tostring_custom_format
$t = [System.TimeOnly]::new(8, 5, 9)
$str = $t.ToString("HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
if ($str -ne "08:05:09") { Write-Host "FAIL: TimeOnly custom format failed, got $str"; exit 1 }
Write-Host "PASS"; exit 0
