# vybe-test: powershell/string_substitution_rules/substitute_in_heredoc
$tag = 'alpha'
$body = @"
value=$tag
"@
if ($body -notmatch 'value=alpha') {
    Write-Host "FAIL: double-quoted here-string did not interpolate variable: $body"
    exit 1
}

Write-Host 'PASS'
exit 0
