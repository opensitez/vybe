# vybe-test: powershell/string_substitution_rules/substitute_escape_dollar
$result = "price is `$100"
if ($result -ne 'price is $100') {
    Write-Host "FAIL: escaped dollar should remain literal: '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
