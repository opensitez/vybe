# vybe-test: powershell/string_substitution_rules/substitute_boolean_text
$isOn = $true
$result = "$isOn"
if ($result -ne 'True') {
    Write-Host "FAIL: boolean substitution expected 'True', got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
