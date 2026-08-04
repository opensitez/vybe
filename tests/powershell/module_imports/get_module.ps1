# vybe-test: powershell/module_imports/get_module
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
if ((Get-Module -Name Microsoft.PowerShell.Utility).Name -eq 'Microsoft.PowerShell.Utility') { exit 0 }
exit 1
