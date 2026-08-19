# vybe-test: powershell/lexical_escape_rules/escape_colon
$val = "http`://example"
if ($val -ne 'http://example') {
    Write-Host "FAIL: colon escape behavior mismatch: $val"
    exit 1
}

Write-Host 'PASS'
exit 0
