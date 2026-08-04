# vybe-test: powershell/module_imports/import_module_alias
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
if ((Get-Command Get-Date).ModuleName -eq 'Microsoft.PowerShell.Utility') { exit 0 }
exit 1
