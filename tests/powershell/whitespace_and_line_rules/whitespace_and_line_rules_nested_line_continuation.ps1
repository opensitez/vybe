# vybe-test: powershell/whitespace_and_line_rules/nested_line_continuation
$sum = 1 + `
    2 + `
    3

if ($sum -ne 6) {
    Write-Host "FAIL: nested continuation returned $sum"
    exit 1
}

Write-Host 'PASS'
exit 0
