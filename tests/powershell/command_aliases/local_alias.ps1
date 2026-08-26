# vybe-test: powershell/command_aliases/local_alias
Set-Alias -Name "lalias" -Value "Get-Date" -Scope Local
$target = (Get-Alias -Name "lalias").Definition
Remove-Item alias:lalias -Force
if ($target -ne "Get-Date") {
    Write-Host "FAIL: Local alias failed"
    exit 1
}
Write-Host "PASS"
exit 0
