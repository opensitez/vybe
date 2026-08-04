# vybe-test: powershell/variable_modifiers/readonly_read
Set-Variable -Name x -Value 1 -Option ReadOnly
if ($x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
