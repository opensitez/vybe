# vybe-test: powershell/readonly_variables/readonly_variable_reassignment_fails
New-Variable -Name "PROTECTED_VAR" -Value 10 -Option ReadOnly
try {
    $PROTECTED_VAR = 20
    Write-Host "FAIL: Reassignment to ReadOnly variable succeeded without -Force, expected throw"
    exit 1
} catch {
    if ($PROTECTED_VAR -ne 10) {
        Write-Host "FAIL: ReadOnly variable mutated despite throw"
        exit 1
    }
}
Write-Host "PASS"
exit 0
