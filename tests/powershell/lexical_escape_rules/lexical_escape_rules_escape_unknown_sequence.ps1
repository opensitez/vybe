# vybe-test: powershell/lexical_escape_rules/escape_unknown_sequence
$val = "a`q"
if ($val -ne 'aq') {
    Write-Host "FAIL: unknown backtick sequence did not treat as escaped char literal: $val"
    exit 1
}

Write-Host 'PASS'
exit 0
