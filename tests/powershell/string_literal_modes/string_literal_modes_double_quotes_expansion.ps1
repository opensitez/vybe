# vybe-test: powershell/string_literal_modes/double_quotes_expansion
$name = 'World'
$result = "Hello $name"
if ($result -ne 'Hello World') {
    Write-Host "FAIL: expected expanded text, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
