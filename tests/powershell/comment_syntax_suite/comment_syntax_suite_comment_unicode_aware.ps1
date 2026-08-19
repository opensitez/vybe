# vybe-test: powershell/comment_syntax_suite/comment_unicode_aware
$text = "value"
# 注释里有中文字符和符号：😀
if ($text -ne 'value') {
    Write-Host 'FAIL: unicode-commented line disturbed variable value'
    exit 1
}

Write-Host 'PASS'
exit 0
