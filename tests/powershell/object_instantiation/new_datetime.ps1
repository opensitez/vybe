# vybe-test: powershell/object_instantiation/new_datetime
$date = Get-Date
if ($date -eq $null) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
