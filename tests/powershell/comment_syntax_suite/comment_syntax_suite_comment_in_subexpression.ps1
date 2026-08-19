# vybe-test: powershell/comment_syntax_suite/comment_in_subexpression
$value = (1 + ( # inline comment in subexpr
2 ))
if ($value -ne 3) {
    Write-Host "FAIL: subexpression with comment malformed: $value"
    exit 1
}

Write-Host 'PASS'
exit 0
