# vybe-test: powershell/whitespace_and_line_rules/unicode_whitespace_boundaries
$u = [char]0x2003
$code = "10${u}+${u}5"
if ((Invoke-Expression $code) -ne 15) {
    Write-Host 'FAIL: unicode whitespace did not behave as token boundary around operator'
    exit 1
}

Write-Host 'PASS'
exit 0
