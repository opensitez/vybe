# vybe-test: powershell/command_aliases/alias_define
Set-Alias -Name "talias" -Value "Get-Date" -Scope Local
$target = (Get-Alias -Name "talias").Definition
Remove-Item alias:talias -Force
if ($target -ne "Get-Date") {
    Write-Host "FAIL: Alias define failed"
    exit 1
}
Write-Host "PASS"
exit 0
