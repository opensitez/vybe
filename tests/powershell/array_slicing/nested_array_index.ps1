# vybe-test: powershell/array_slicing/nested_array_index
$arr = @(1, @(2,3), 4)
if ($arr[1][0] -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
