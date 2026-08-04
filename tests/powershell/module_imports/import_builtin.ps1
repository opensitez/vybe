# vybe-test: powershell/module_imports/import_builtin
Import-Module Microsoft.PowerShell.Management -ErrorAction Stop
if ((Get-Command Get-Item).ModuleName -eq 'Microsoft.PowerShell.Management') { exit 0 }
exit 1
