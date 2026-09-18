# vybe-test: powershell/invoke_expression_cmdlet/iex_string_variable_interpolation_inside_evaluated_string
# Invoke-Expression expands variables from the calling scope inside double-quoted evaluated strings
$greeting = "World"
$result = Invoke-Expression '"Hello, $greeting!"'

if ($result -ne "Hello, World!") {
    Write-Host "FAIL: interpolation mismatch, got: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
