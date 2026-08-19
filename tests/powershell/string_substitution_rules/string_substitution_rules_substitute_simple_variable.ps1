# vybe-test: powershell/string_substitution_rules/substitute_simple_variable
$val = 'alpha'
$result = "$val"
if ($result -ne 'alpha') {
    Write-Host "FAIL: simple variable substitution expected alpha, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
