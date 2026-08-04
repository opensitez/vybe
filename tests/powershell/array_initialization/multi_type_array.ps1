# vybe-test: powershell/array_initialization/multi_type_array
$arr = 1,'A',$true
if ($arr.Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
