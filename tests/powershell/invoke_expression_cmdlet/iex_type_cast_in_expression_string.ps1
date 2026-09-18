# vybe-test: powershell/invoke_expression_cmdlet/iex_type_cast_in_expression_string
# Invoke-Expression evaluates type cast syntax in the command string
$result = Invoke-Expression "[int]'42'"

if ($result -ne 42 -or $result.GetType().Name -ne "Int32") {
    Write-Host "FAIL: expected Int32 42, got: $result ($($result.GetType().Name))"
    exit 1
}

Write-Host "PASS"
exit 0
