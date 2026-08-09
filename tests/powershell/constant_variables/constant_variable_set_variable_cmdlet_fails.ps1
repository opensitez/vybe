# vybe-test: powershell/constant_variables/constant_variable_set_variable_cmdlet_fails
New-Variable -Name "CMDLET_CONST" -Value "Immutable" -Option Constant
try {
    Set-Variable -Name "CMDLET_CONST" -Value "Changed" -ErrorAction Stop
    Write-Host "FAIL: Set-Variable on Constant variable succeeded, expected throw"
    exit 1
} catch {
    if ($CMDLET_CONST -ne "Immutable") {
        Write-Host "FAIL: Set-Variable modified Constant variable"
        exit 1
    }
}
Write-Host "PASS"
exit 0
