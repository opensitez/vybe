# vybe-test: powershell/array_initialization/nested_array
$arr = 1,@(2,3)
if ($arr[1][0] -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
