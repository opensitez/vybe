# vybe-test: powershell/invoke_expression_cmdlet/iex_boolean_expression_returns_bool
# Invoke-Expression can evaluate comparison operators and return a Boolean value
$result = Invoke-Expression '3 -gt 2'

if ($result -ne $true) {
    Write-Host "FAIL: expected `$true from '3 -gt 2', got: $result"
    exit 1
}

Write-Host "PASS"
exit 0
