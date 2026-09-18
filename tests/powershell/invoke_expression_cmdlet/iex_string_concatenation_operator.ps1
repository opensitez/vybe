# vybe-test: powershell/invoke_expression_cmdlet/iex_string_concatenation_operator
# Invoke-Expression evaluates string concatenation using the + operator inside the command string
$result = Invoke-Expression '"foo" + "bar"'

if ($result -ne "foobar") {
    Write-Host "FAIL: expected 'foobar', got: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
