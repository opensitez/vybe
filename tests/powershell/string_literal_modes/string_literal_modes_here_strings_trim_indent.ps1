# vybe-test: powershell/string_literal_modes/here_strings_trim_indent
$indented = @"
    indented one
    indented two
"@

if ($indented -notmatch '^[ ]{4}indented one`r?`n[ ]{4}indented two`r?`n$') {
    Write-Host 'FAIL: here-string indentation should be preserved with leading spaces on each line'
    exit 1
}

Write-Host 'PASS'
exit 0
