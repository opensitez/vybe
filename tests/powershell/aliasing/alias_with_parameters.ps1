# vybe-test: powershell/aliasing/alias_with_parameters
Set-Alias hi Write-Output
if ((hi 'x') -ne 'x') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
