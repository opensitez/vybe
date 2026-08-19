# vybe-test: powershell/string_substitution_rules/substitute_primitive_object
$num = 42
$result = "$num"
if ($result -ne '42') {
    Write-Host "FAIL: primitive number substitution expected 42, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
