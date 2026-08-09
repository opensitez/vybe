# vybe-test: powershell/readonly_variables/readonly_variable_set
New-Variable -Name "MY_READONLY" -Value "ReadOnlyData" -Option ReadOnly
if ($MY_READONLY -ne "ReadOnlyData") {
    Write-Host "FAIL: ReadOnly variable set expected ReadOnlyData, got $MY_READONLY"
    exit 1
}
Write-Host "PASS"
exit 0
