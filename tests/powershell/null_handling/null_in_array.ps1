# vybe-test: powershell/null_handling/null_in_array
$arr = 1,$null,3
if ($arr[1] -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
