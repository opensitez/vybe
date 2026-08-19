# vybe-test: powershell/string_substitution_rules/substitute_subexpression
$result = "$( (2 + 3) * 2 )"
if ($result -ne '10') {
    Write-Host "FAIL: subexpression substitution expected 10, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
