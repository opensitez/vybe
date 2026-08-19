# vybe-test: powershell/whitespace_and_line_rules/script_block_trailing_backtick
$fn = & {
    9 + `
    1
}

if ($fn -ne 10) {
    Write-Host "FAIL: trailing backtick in script block produced $fn"
    exit 1
}

Write-Host 'PASS'
exit 0
