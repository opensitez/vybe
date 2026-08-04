# vybe-test: powershell/comments/comment_in_function
function Test-Func {
    # inside function
    Write-Output 'PASS'
}
Test-Func
exit 0
