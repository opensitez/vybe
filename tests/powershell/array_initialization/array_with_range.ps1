# vybe-test: powershell/array_initialization/array_with_range
$arr = 1..3
if ($arr[2] -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
