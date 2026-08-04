# vybe-test: powershell/dot_sourcing/source_script
$script = "$PWD/dot_sourcing_script.ps1"
Set-Content -Path $script -Value 'Write-Output "SOURCED"'
. $script | Out-Null
Remove-Item $script
Write-Host 'PASS'
exit 0
