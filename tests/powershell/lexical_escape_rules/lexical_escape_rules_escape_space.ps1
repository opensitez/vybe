# vybe-test: powershell/lexical_escape_rules/escape_space
$val = "a` b"
if ($val -ne 'a b') {
    Write-Host "FAIL: escaped space should remain literal space in string, got '$val'"
    exit 1
}

Write-Host 'PASS'
exit 0
