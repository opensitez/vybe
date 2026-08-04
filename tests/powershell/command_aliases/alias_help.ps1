# vybe-test: powershell/command_aliases/alias_help
Set-Alias inf Test-Path
if ((Get-Alias inf).Name -eq 'inf') { exit 0 }
exit 1
