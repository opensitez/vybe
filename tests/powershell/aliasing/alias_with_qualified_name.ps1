# vybe-test: powershell/aliasing/alias_with_qualified_name
Set-Alias hi Write-Output
if ((hi 'PASS') -ne 'PASS') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
