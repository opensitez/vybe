# vybe-test: powershell/command_aliases/alias_with_parameters
Set-Alias say Write-Output
if ((say 'param') -eq 'param') { exit 0 }
exit 1
