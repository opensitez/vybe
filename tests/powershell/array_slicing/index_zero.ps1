# vybe-test: powershell/array_slicing/index_zero
$arr = 1,2,3
if ($arr[0] -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
