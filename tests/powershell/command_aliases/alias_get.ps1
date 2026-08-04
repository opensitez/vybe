# vybe-test: powershell/command_aliases/alias_get
Set-Alias testalias Write-Host
if ((Get-Alias testalias).Definition -eq 'Write-Host') { exit 0 }
exit 1
