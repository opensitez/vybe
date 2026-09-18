# vybe-test: powershell/variables/clear_variable_resets_value_to_null
# Clear-Variable resets the variable's value to null while leaving the variable defined
$testVar = "initial value"

Clear-Variable -Name "testVar"

if ($testVar -ne $null) {
    Write-Host "FAIL: variable was not cleared to null, got '$testVar'"
    exit 1
}

# The variable itself should still exist in the session
$varObj = Get-Variable -Name "testVar" -ErrorAction SilentlyContinue
if ($varObj -eq $null) {
    Write-Host "FAIL: variable object was deleted instead of cleared"
    exit 1
}

Write-Host "PASS"
exit 0
