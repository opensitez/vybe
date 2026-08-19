# vybe-test: powershell/comment_syntax_suite/comment_preserves_string_hash
$snippet = '#not-a-comment'
if ($snippet -ne '#not-a-comment') {
    Write-Host "FAIL: single-quoted hash text changed: $snippet"
    exit 1
}

Write-Host 'PASS'
exit 0
