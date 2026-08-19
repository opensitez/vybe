# vybe-test: powershell/string_literal_modes/single_quotes_verbatim_text
$single = 'path: C:\tmp\`nliteral\`$name'

if (-not $single.Contains('`n')) {
    Write-Host 'FAIL: single-quoted string should keep backtick-n literally'
    exit 1
}

if (-not $single.Contains('`$name')) {
    Write-Host 'FAIL: single-quoted string should keep dollar text literally'
    exit 1
}

Write-Host 'PASS'
exit 0
