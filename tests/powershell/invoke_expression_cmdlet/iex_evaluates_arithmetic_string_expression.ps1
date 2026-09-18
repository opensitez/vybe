# vybe-test: powershell/invoke_expression_cmdlet/iex_evaluates_arithmetic_string_expression
# Invoke-Expression evaluates a string containing a PowerShell arithmetic expression and returns the result
$result = Invoke-Expression "2 + 2"

if ($result -ne 4) {
    Write-Host "FAIL: expected 4, got: $result"
    exit 1
}

Write-Host "PASS"
exit 0
