# vybe-test: powershell/lexical_escape_rules/escape_semicolon
$val = "a`;b"
if ($val -ne 'a;b') {
    Write-Host "FAIL: semicolon escaped in string should remain literal, got $val"
    exit 1
}

Write-Host 'PASS'
exit 0
