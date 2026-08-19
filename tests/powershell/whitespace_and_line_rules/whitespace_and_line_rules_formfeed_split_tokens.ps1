# vybe-test: powershell/whitespace_and_line_rules/formfeed_split_tokens
$ff = [char]0x000C
$expr = "8${ff}+${ff}7"
if ((Invoke-Expression $expr) -ne 15) {
    Write-Host 'FAIL: form feed was not treated as token separator'
    exit 1
}

Write-Host 'PASS'
exit 0
