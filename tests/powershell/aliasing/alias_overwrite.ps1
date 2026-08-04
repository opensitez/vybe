# vybe-test: powershell/aliasing/alias_overwrite
Set-Alias hi Write-Output
Set-Alias hi Echo -Force
if ((hi 'PASS') -ne 'PASS') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
