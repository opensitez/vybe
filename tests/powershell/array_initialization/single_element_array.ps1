# vybe-test: powershell/array_initialization/single_element_array
$arr = ,1
if ($arr.Count -eq 1 -and $arr[0] -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
