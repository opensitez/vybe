# vybe-test: powershell/invoke_expression_cmdlet/iex_multi_statement_returns_last_value
# Invoke-Expression evaluates multiple statements and returns the result of the final one
$result = Invoke-Expression '$x = 10; $x * 3'

if ($result -ne 30) {
    Write-Host "FAIL: expected 30 from multi-statement expression, got: $result"
    exit 1
}

Write-Host "PASS"
exit 0
