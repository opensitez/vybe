# vybe-test: powershell/module_imports/module_scope
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
if ((Get-Command Get-Unique).ModuleName -eq 'Microsoft.PowerShell.Utility') { exit 0 }
exit 1
