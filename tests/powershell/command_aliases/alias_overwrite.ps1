# vybe-test: powershell/command_aliases/alias_overwrite
Set-Alias now Write-Host
Set-Alias now Write-Output -Force
if ((now 'a') -eq 'a') { exit 0 }
exit 1
