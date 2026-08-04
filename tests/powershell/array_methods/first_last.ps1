# vybe-test: powershell/array_methods/first_last
$arr = 1,2,3
if ($arr[0] -eq 1 -and $arr[-1] -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
