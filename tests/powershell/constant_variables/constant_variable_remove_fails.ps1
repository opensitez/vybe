# vybe-test: powershell/constant_variables/constant_variable_remove_fails
New-Variable -Name "PERM_CONST" -Value "Fixed" -Option Constant
try {
    Remove-Variable -Name "PERM_CONST" -ErrorAction Stop
    Write-Host "FAIL: Remove-Variable on Constant variable succeeded, expected throw"
    exit 1
} catch {
    if ($PERM_CONST -ne "Fixed") {
        Write-Host "FAIL: PERM_CONST value missing after failed removal"
        exit 1
    }
}
Write-Host "PASS"
exit 0
