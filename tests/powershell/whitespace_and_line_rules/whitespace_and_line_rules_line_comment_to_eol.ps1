# vybe-test: powershell/whitespace_and_line_rules/line_comment_to_eol
$x = 4 + 5 # comment until EOL
if ($x -ne 9) {
    Write-Host "FAIL: inline line comment malformed, got $x"
    exit 1
}

Write-Host 'PASS'
exit 0
