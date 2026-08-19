# vybe-test: powershell/whitespace_and_line_rules/crlf_preserved_in_here_strings
$here = @"
alpha`r`nbeta
"@

if (-not $here.Contains("`r`n")) {
    Write-Host 'FAIL: expected explicit CRLF sequence in here-string'
    exit 1
}

if ($here.Trim("`r", "`n").Length -ne 9) {
    Write-Host 'FAIL: here-string content corrupted'
    exit 1
}

Write-Host 'PASS'
exit 0
