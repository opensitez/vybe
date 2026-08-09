# vybe-test: powershell/readonly_variables/readonly_variable_force_clear
New-Variable -Name "FORCE_CLEAR_RO" -Value "Data" -Option ReadOnly
Clear-Variable -Name "FORCE_CLEAR_RO" -Force
if ($FORCE_CLEAR_RO -ne $null) {
    Write-Host "FAIL: Clear-Variable -Force on ReadOnly variable expected null, got $FORCE_CLEAR_RO"
    exit 1
}
Write-Host "PASS"
exit 0
