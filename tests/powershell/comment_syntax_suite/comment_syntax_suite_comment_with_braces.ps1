# vybe-test: powershell/comment_syntax_suite/comment_with_braces
# comment contains { } and "quotes" that should not form a block
$dict = @{ key = 1 }
if ($dict.key -ne 1) {
    Write-Host 'FAIL: braces in comment altered parser blocks'
    exit 1
}

Write-Host 'PASS'
exit 0
