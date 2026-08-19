# vybe-test: powershell/string_substitution_rules/substitute_dollar_parenthesized
$left = 1
$right = 2
$result = "$( $left + $right )"
if ($result -ne '3') {
    Write-Host "FAIL: parenthesized substitution expected 3, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
