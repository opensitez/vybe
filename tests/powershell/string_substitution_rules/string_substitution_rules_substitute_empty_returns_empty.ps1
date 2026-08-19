# vybe-test: powershell/string_substitution_rules/substitute_empty_returns_empty
$empty = ''
$result = "$empty"
if ($result.Length -ne 0) {
    Write-Host "FAIL: empty string in substitution expected zero length, got $($result.Length)"
    exit 1
}

Write-Host 'PASS'
exit 0
