# vybe-test: powershell/command_aliases/global_alias
Set-Alias -Name "galias" -Value "Get-Date" -Scope Global
$target = (Get-Alias -Name "galias").Definition
Remove-Item alias:galias -Force
if ($target -ne "Get-Date") {
    Write-Host "FAIL: Global alias failed"
    exit 1
}
Write-Host "PASS"
exit 0
