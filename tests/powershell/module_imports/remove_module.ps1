# vybe-test: powershell/module_imports/remove_module
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
Remove-Module Microsoft.PowerShell.Utility -ErrorAction SilentlyContinue
if ((Get-Module Microsoft.PowerShell.Utility) -eq $null) { exit 0 }
exit 1
