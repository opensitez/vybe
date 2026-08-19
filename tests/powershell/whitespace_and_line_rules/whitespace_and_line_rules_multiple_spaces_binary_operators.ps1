# vybe-test: powershell/whitespace_and_line_rules/multiple_spaces_binary_operators
$sum = 1    +   2    *    3
if ($sum -ne 7) {
    Write-Host "FAIL: expected 7 with mixed spaces around operators, got $sum"
    exit 1
}

Write-Host 'PASS'
exit 0
