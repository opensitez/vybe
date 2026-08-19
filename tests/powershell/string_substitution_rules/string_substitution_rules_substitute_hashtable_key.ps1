# vybe-test: powershell/string_substitution_rules/substitute_hashtable_key
$map = @{ kind = 'value' }
$result = "$($map['kind'])"
if ($result -ne 'value') {
    Write-Host "FAIL: expected hashtable value, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
