# vybe-test: powershell/whitespace_and_line_rules/whitespace_and_line_rules_crlf_preserved_in_here_strings
$str = "Line1`n`tLine2`$val`"quote`""
if ($str.Length -gt 0) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
