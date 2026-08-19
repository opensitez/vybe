# vybe-test: powershell/comment_syntax_suite/multiline_comment_closes
<#
    start
#> $value = 12
if ($value -ne 12) {
    Write-Host "FAIL: comment close handling failed"
    exit 1
}

Write-Host 'PASS'
exit 0
