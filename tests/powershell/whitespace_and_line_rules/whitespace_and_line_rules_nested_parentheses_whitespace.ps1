# vybe-test: powershell/whitespace_and_line_rules/nested_parentheses_whitespace
$result = (((2 + 3))   + ((1 + 1)))

if ($result -ne 7) {
    Write-Host "FAIL: nested whitespaceed parentheses failed, got $result"
    exit 1
}

Write-Host 'PASS'
exit 0
