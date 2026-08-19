# vybe-test: powershell/string_substitution_rules/substitute_automatic_variable
$result = "$PID"
if ($result -match '^\d+$' -ne $true) {
    Write-Host "FAIL: automatic variable PID should produce a numeric string, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
