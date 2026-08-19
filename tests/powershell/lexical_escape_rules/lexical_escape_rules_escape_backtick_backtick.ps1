# vybe-test: powershell/lexical_escape_rules/escape_backtick_backtick
$value = "a``b"
if ($value -ne 'a`b') {
    Write-Host "FAIL: expected literal backtick, got $value"
    exit 1
}

Write-Host 'PASS'
exit 0
