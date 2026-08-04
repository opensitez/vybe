# vybe-test: powershell/variable_modifiers/allscope
Set-Variable -Name x -Value 1 -Option AllScope
if ($x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
