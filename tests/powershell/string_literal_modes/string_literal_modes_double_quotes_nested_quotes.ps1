# vybe-test: powershell/string_literal_modes/double_quotes_nested_quotes
$result = "He said `"Hi`""
if ($result -ne 'He said "Hi"') {
    Write-Host "FAIL: expected quoted nested string, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
