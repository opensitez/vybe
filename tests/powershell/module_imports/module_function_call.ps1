# vybe-test: powershell/module_imports/module_function_call
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
if ((Get-Date -Date '2000-01-01').Year -eq 2000) { exit 0 }
exit 1
