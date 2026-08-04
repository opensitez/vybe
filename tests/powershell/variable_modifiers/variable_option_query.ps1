# vybe-test: powershell/variable_modifiers/variable_option_query
Set-Variable -Name x -Value 1 -Option ReadOnly
if ((Get-Variable -Name x).Options -band [System.Management.Automation.ScopeDescriptionOptions]::ReadOnly) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
