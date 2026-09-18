# vybe-test: powershell/invoke_expression_cmdlet/iex_conditional_expression_if_else
# Invoke-Expression evaluates an if/else conditional expression and returns the selected branch value
$result = Invoke-Expression 'if ($true) { "yes" } else { "no" }'

if ($result -ne "yes") {
    Write-Host "FAIL: expected 'yes' from conditional expression, got: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
