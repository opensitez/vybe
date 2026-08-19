# vybe-test: powershell/string_substitution_rules/substitute_split_lines
$line = "left`nright"
if ($line -notmatch '^left`nright$') {
    Write-Host "FAIL: expected embedded newline before substitution, got '$line'"
    exit 1
}

Write-Host 'PASS'
exit 0
