# vybe-test: powershell/variable_declaration/variable_from_command
$x = Get-Date
if ($x -eq $null) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
