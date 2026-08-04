# vybe-test: powershell/command_aliases/alias_define
Set-Alias greet Write-Host
if ((greet 'hello') -eq 'hello') { exit 0 }
exit 1
