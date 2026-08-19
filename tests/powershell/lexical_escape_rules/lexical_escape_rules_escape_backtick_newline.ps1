# vybe-test: powershell/lexical_escape_rules/escape_backtick_newline
$result = 1 + `
    2
if ($result -ne 3) {
    Write-Host "FAIL: backtick newline continuation produced $result"
    exit 1
}

Write-Host 'PASS'
exit 0
