# vybe-test: powershell/comment_syntax_suite/comment_inside_function_body
function comment_syntax_suite_comment_inside_function_body {
    # internal comment does not execute
    return 17
}

if ((comment_syntax_suite_comment_inside_function_body) -ne 17) {
    Write-Host 'FAIL: comment inside function body changed behavior'
    exit 1
}

Write-Host 'PASS'
exit 0
