# vybe-test: powershell/comment_syntax_suite/comment_trailing_after_pipe
$result = 1, 2, 3 |
    Where-Object { $_ -ge 2 } # comment after pipe continuation line

if (($result -join ',') -ne '2,3') {
    Write-Host "FAIL: comment after pipeline changed output: $($result -join ',')"
    exit 1
}

Write-Host 'PASS'
exit 0
