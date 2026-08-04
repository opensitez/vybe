# vybe-test: powershell/array_slicing/assignment_index
$arr = 1,2,3
$arr[1] = 5
if ($arr[1] -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
