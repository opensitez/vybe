# vybe-test: powershell/token_and_grammar_system/line_comment_mid_expression
$value = 2 + # inline ignore
1
if ($value -ne 3) {
    Write-Host "FAIL: line comment mid-expression failed, value=$value"
    exit 1
}

Write-Host 'PASS'
exit 0
