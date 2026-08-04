# vybe-test: powershell/module_imports/list_modules
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
if ((Get-Module).Count -ge 1) { exit 0 }
exit 1
