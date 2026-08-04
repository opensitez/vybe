# vybe-test: powershell/aliasing/alias_in_function
function Test-Func {
    Set-Alias hi Write-Output
}
Test-Func
Write-Host 'PASS'
exit 0
