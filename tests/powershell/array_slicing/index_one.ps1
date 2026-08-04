# vybe-test: powershell/array_slicing/index_one
$arr = 1,2,3
if ($arr[1] -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
