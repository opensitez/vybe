# vybe-test: powershell/command_aliases/alias_remove
Set-Alias temp Write-Host
Remove-Item Alias:\temp
if ((Get-Command temp -ErrorAction SilentlyContinue) -eq $null) { exit 0 }
exit 1
