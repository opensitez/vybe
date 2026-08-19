# vybe-test: powershell/lexical_escape_rules/escape_curly_braces
$val = "a`{x` }"
if ($val -ne 'a{x }') {
    Write-Host "FAIL: escaped braces not preserved literally: $val"
    exit 1
}

Write-Host 'PASS'
exit 0
