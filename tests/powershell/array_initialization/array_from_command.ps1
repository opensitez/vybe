# vybe-test: powershell/array_initialization/array_from_command
$arr = @(1,2,3 | ForEach-Object { $_ })
if ($arr.Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
