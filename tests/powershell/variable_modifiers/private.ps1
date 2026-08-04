# vybe-test: powershell/variable_modifiers/private
Set-Variable -Name x -Value 1 -Option Private
if ($x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
