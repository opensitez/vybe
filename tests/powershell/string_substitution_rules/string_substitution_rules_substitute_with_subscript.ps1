# vybe-test: powershell/string_substitution_rules/substitute_with_subscript
$payload = @{
    nested = @('zero', 'one', 'two')
}
$result = "$($payload.nested[1])"
if ($result -ne 'one') {
    Write-Host "FAIL: subscript substitution expected one, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
