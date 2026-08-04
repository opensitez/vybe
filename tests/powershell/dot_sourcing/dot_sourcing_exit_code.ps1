# vybe-test: powershell/dot_sourcing/dot_sourcing_exit_code
$script = "$PWD/dot_sourcing_exit.ps1"
Set-Content -Path $script -Value 'exit 0'
. $script
Remove-Item $script
Write-Host 'PASS'
exit 0
