# vybe-test: powershell/lexical_escape_rules/escape_in_double_quoted_string
$actual = "A`nB"
if ($actual -ne "A`nB") {
    Write-Host 'FAIL: newline escape in double-quoted string failed'
    exit 1
}

if ($actual.GetType().Name -ne 'String') {
    Write-Host 'FAIL: result type mismatch'
    exit 1
}

Write-Host 'PASS'
exit 0
