# vybe-test: powershell/array_methods/map
$arr = 1,2,3 | ForEach-Object { $_ * 2 }
if (($arr -join ',') -eq '2,4,6') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
