# vybe-test: powershell/aliasing/alias_help
Set-Alias hi Write-Output
Get-Help hi | Out-Null
Write-Host 'PASS'
exit 0
