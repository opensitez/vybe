# vybe-test: powershell/dot_sourcing/dot_sourcing_command
$script = "$PWD/dot_sourcing_cmd.ps1"
Set-Content -Path $script -Value 'Get-Command Write-Host | Out-Null'
. $script
Remove-Item $script
Write-Host 'PASS'
exit 0
