# vybe-test: powershell/array_methods/sort
$arr = 3,1,2
$arr = $arr | Sort-Object
if ($arr[0] -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
