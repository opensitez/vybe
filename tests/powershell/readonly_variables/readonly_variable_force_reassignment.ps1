# vybe-test: powershell/readonly_variables/readonly_variable_force_reassignment
New-Variable -Name "OVERRIDABLE_RO" -Value "Initial" -Option ReadOnly
Set-Variable -Name "OVERRIDABLE_RO" -Value "ForcedValue" -Force
if ($OVERRIDABLE_RO -ne "ForcedValue") {
    Write-Host "FAIL: Set-Variable -Force on ReadOnly variable expected ForcedValue, got $OVERRIDABLE_RO"
    exit 1
}
Write-Host "PASS"
exit 0
