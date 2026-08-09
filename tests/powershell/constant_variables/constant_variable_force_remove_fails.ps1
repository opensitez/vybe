# vybe-test: powershell/constant_variables/constant_variable_force_remove_fails
New-Variable -Name "FORCE_CONST" -Value "Locked" -Option Constant
try {
    Remove-Variable -Name "FORCE_CONST" -Force -ErrorAction Stop
    Write-Host "FAIL: Remove-Variable -Force on Constant variable succeeded, expected throw"
    exit 1
} catch {
    if ($FORCE_CONST -ne "Locked") {
        Write-Host "FAIL: Constant variable removed despite -Force error"
        exit 1
    }
}
Write-Host "PASS"
exit 0
