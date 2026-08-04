# vybe-test: powershell/aliasing/set_alias
Set-Alias hi Write-Output
if ((hi 'PASS') -ne 'PASS') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
