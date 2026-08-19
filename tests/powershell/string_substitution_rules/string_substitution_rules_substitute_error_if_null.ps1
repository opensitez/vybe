# vybe-test: powershell/string_substitution_rules/substitute_error_if_null
$maybe = $null
$result = "$maybe"
if ($result -ne '') {
    Write-Host "FAIL: null substitution expected empty text, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
