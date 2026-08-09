# vybe-test: powershell/readonly_variables/readonly_variable_force_remove
New-Variable -Name "CAN_FORCE_REM" -Value "Temp" -Option ReadOnly
Remove-Variable -Name "CAN_FORCE_REM" -Force
if (Test-Path "variable:CAN_FORCE_REM") {
    Write-Host "FAIL: Remove-Variable -Force failed to remove ReadOnly variable"
    exit 1
}
Write-Host "PASS"
exit 0
