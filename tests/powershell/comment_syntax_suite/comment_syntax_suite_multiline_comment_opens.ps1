# vybe-test: powershell/comment_syntax_suite/multiline_comment_opens
<# outer
   nested open-like text
#>
$value = 11
if ($value -ne 11) {
    Write-Host "FAIL: multiline comment open/close should not affect value"
    exit 1
}

Write-Host 'PASS'
exit 0
