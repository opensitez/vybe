# vybe-test: powershell/comment_syntax_suite/comment_with_url_text
$text = 'keep https://example.com/path?x=1#frag in string'
$text2 = "value=$text"
if ($text2 -notlike '*https://example.com/path?x=1#frag*') {
    Write-Host 'FAIL: URL text in comment-like syntax should stay untouched in string'
    exit 1
}

Write-Host 'PASS'
exit 0
