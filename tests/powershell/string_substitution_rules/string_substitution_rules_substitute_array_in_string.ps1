# vybe-test: powershell/string_substitution_rules/substitute_array_in_string
$vals = @(7, 8, 9)
$result = "$($vals[1])"
if ($result -ne '8') {
    Write-Host "FAIL: array item substitution expected 8, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
