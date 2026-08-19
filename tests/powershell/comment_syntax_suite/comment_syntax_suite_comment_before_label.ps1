# vybe-test: powershell/comment_syntax_suite/comment_before_label
# comment before first label should be ignored by parser
:label_check
$hit = 1

# comment before second label as well
:label_two
$hit += 1

if ($hit -ne 2) {
    Write-Host "FAIL: labels with comments nearby changed execution"
    exit 1
}

Write-Host 'PASS'
exit 0
