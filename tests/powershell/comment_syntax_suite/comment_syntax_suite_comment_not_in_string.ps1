# vybe-test: powershell/comment_syntax_suite/comment_not_in_string
$raw = '#this is data'
$quoted = '#not-quoted'
if ($raw -ne '#this is data' -or $quoted -ne '#not-quoted') {
    Write-Host 'FAIL: hash in quoted string became comment or changed'
    exit 1
}

Write-Host 'PASS'
exit 0
