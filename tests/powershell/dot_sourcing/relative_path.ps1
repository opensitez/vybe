# vybe-test: powershell/dot_sourcing/relative_path
Push-Location $PWD
$script = 'dot_sourcing_rel.ps1'
Set-Content -Path $script -Value 'Write-Output "REL"'
. $script | Out-Null
Remove-Item $script
Pop-Location
Write-Host 'PASS'
exit 0
