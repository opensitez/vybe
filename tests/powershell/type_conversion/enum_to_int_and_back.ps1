# vybe-test: powershell/type_conversion/enum_to_int_and_back
$day = [DayOfWeek]::Wednesday
$num = [int]$day
if ($num -ne 3) { Write-Host "FAIL: Wednesday should be 3, got $num"; exit 1 }
$back = [DayOfWeek]$num
if ($back -ne [DayOfWeek]::Wednesday) { Write-Host "FAIL: int 3 back to enum"; exit 1 }
Write-Host "PASS"
exit 0
