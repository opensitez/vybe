# vybe-test: powershell/lexical_escape_rules/escape_not_in_double_quotes
$actual = 'x`q'
if ($actual -ne 'x`q') {
    Write-Host "FAIL: single-quoted unknown escape should keep literal text, got $actual"
    exit 1
}

Write-Host 'PASS'
exit 0
