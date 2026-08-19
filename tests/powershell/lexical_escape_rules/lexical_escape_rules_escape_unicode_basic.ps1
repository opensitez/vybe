# vybe-test: powershell/lexical_escape_rules/escape_unicode_basic
$pi = "\u{03A0}"
if ($pi -ne 'Π') {
    Write-Host "FAIL: expected Greek Pi, got $pi"
    exit 1
}

Write-Host 'PASS'
exit 0
