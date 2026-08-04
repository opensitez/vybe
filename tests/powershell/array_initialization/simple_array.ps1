# vybe-test: powershell/array_initialization/simple_array
$arr = 1,2,3
if ($arr[0] -eq 1 -and $arr.Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
