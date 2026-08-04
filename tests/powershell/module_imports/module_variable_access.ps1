# vybe-test: powershell/module_imports/module_variable_access
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
if ((Get-Module Microsoft.PowerShell.Utility).Name -eq 'Microsoft.PowerShell.Utility') { exit 0 }
exit 1
