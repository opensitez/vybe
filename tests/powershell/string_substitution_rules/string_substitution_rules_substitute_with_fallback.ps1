# vybe-test: powershell/string_substitution_rules/substitute_with_fallback
$name = $null
$value = if ($null -eq $name) { 'fallback' } else { $name }
$result = "$value"
if ($result -ne 'fallback') {
    Write-Host "FAIL: fallback substitution expected 'fallback', got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
