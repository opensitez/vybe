# vybe-test: powershell/invoke_expression_cmdlet/iex_variable_assignment_visible_in_caller_scope
# Variables assigned inside Invoke-Expression are visible in the calling scope
Invoke-Expression '$callerScopeVar = "visible"'

if ($callerScopeVar -ne "visible") {
    Write-Host "FAIL: expected 'visible' in caller scope, got: '$callerScopeVar'"
    exit 1
}

Write-Host "PASS"
exit 0
