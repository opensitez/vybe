# vybe-test: powershell/lexical_escape_rules/escape_backtick_quote
$val = "\""
if ($val -ne '"') {
    Write-Host "FAIL: expected quoted literal, got $val"
    exit 1
}

Write-Host 'PASS'
exit 0
