# vybe-test: powershell/string_literal_modes/single_quotes_no_expansion
$name = 'World'
$result = 'Hello $name'
if ($result -ne 'Hello $name') {
    Write-Host "FAIL: single-quoted text was unexpectedly expanded: '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
