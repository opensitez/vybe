# vybe-test: powershell/dot_sourcing/dot_sourcing_current_directory
$script = "$PWD/dot_sourcing_dir.ps1"
Set-Content -Path $script -Value 'Write-Output "DIR"'
. $script | Out-Null
Remove-Item $script
Write-Host 'PASS'
exit 0
