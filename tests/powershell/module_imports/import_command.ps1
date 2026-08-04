# vybe-test: powershell/module_imports/import_command
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
if ((Get-Command Measure-Object).Name -eq 'Measure-Object') { exit 0 }
exit 1
