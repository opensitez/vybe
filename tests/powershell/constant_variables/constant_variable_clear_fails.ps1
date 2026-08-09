# vybe-test: powershell/constant_variables/constant_variable_clear_fails
New-Variable -Name "UNCLEAR_CONST" -Value 777 -Option Constant
try {
    Clear-Variable -Name "UNCLEAR_CONST" -ErrorAction Stop
    Write-Host "FAIL: Clear-Variable on Constant variable succeeded, expected throw"
    exit 1
} catch {
    if ($UNCLEAR_CONST -ne 777) {
        Write-Host "FAIL: UNCLEAR_CONST cleared despite error"
        exit 1
    }
}
Write-Host "PASS"
exit 0
