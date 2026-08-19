# vybe-test: powershell/string_substitution_rules/substitute_array_index
$items = @('first', 'second', 'third')
$result = "$($items[2])"
if ($result -ne 'third') {
    Write-Host "FAIL: expected third, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
