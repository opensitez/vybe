# vybe-test: powershell/aliasing/remove_alias
Set-Alias hi Write-Output
Remove-Item Alias:hi
if (Get-Command hi -ErrorAction SilentlyContinue) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
