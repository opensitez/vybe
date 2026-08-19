# vybe-test: powershell/whitespace_and_line_rules/semicolon_as_statement_separator
$a = 1 ; $b = 2 ; $c = $a + $b
if ($c -ne 3) {
    Write-Host "FAIL: semicolon separator failed with $c"
    exit 1
}

Write-Host 'PASS'
exit 0
