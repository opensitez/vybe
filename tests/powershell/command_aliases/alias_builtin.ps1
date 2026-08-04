# vybe-test: powershell/command_aliases/alias_builtin
Set-Alias ls Get-ChildItem
if ((ls . | Get-Member | Where-Object { $_.Name -eq 'Name' }) -ne $null) { exit 0 }
exit 1
