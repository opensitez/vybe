# vybe-test: powershell/lexical_escape_rules/escape_parens
$val = "a`(b`")
if ($val -ne 'a(b)') {
    Write-Host "FAIL: escaped parens should remain literal chars"
    exit 1
}

Write-Host 'PASS'
exit 0
