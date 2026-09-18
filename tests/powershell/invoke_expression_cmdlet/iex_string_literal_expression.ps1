# vybe-test: powershell/invoke_expression_cmdlet/iex_string_literal_expression
# Invoke-Expression evaluates a quoted string expression and returns the string value
$result = Invoke-Expression '"hello world"'

if ($result -ne "hello world") {
    Write-Host "FAIL: expected 'hello world', got: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
