# vybe-test: powershell/array_initialization/empty_array
$arr = @()
if ($arr.Count -eq 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
