# vybe-test: powershell/dot_sourcing/dot_sourcing_includes
$script = "$PWD/dot_sourcing_inc.ps1"
Set-Content -Path $script -Value 'Write-Output "INCLUDE"'
. $script | Out-Null
Remove-Item $script
Write-Host 'PASS'
exit 0
