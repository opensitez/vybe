# vybe-test: powershell/constant_variables/constant_variable_reassignment_fails
New-Variable -Name "IMMUTABLE_VAL" -Value 42 -Option Constant
try {
    $IMMUTABLE_VAL = 999
    Write-Host "FAIL: Reassignment to Constant variable succeeded, expected throw"
    exit 1
} catch {
    if ($IMMUTABLE_VAL -ne 42) {
        Write-Host "FAIL: Constant variable value mutated despite catch"
        exit 1
    }
}
Write-Host "PASS"
exit 0
