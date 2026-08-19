# vybe-test: powershell/lexical_escape_rules/escape_at_sign
$val = "user`@domain"
if ($val -ne 'user@domain') {
    Write-Host "FAIL: escaped at-sign not preserved: $val"
    exit 1
}

Write-Host 'PASS'
exit 0
