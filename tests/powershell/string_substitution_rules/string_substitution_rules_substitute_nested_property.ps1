# vybe-test: powershell/string_substitution_rules/substitute_nested_property
$payload = [pscustomobject]@{
    outer = [pscustomobject]@{ inner = 42 }
}
$result = "$($payload.outer.inner)"
if ($result -ne '42') {
    Write-Host "FAIL: nested property substitution expected 42, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
