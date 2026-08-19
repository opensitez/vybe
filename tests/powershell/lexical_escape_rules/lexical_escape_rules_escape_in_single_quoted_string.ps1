# vybe-test: powershell/lexical_escape_rules/escape_in_single_quoted_string
$actual = 'A`nB'
if ($actual -ne 'A`nB') {
    Write-Host 'FAIL: single-quoted backtick-n should remain literal two-char sequence'
    exit 1
}

Write-Host 'PASS'
exit 0
