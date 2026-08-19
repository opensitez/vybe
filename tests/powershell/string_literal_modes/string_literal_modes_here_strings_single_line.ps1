# vybe-test: powershell/string_literal_modes/here_strings_single_line
$heredoc = @"
single line
"@
if ($heredoc.TrimEnd("`r", "`n") -ne 'single line') {
    Write-Host "FAIL: single-line here-string value incorrect: $heredoc"
    exit 1
}

Write-Host 'PASS'
exit 0
