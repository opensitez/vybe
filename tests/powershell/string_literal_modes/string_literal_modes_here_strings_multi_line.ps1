# vybe-test: powershell/string_literal_modes/here_strings_multi_line
$heredoc = @"
alpha
beta
"@
if ($heredoc -notmatch 'alpha' -or $heredoc -notmatch 'beta') {
    Write-Host "FAIL: multi-line here-string missing expected lines"
    exit 1
}

Write-Host 'PASS'
exit 0
