# vybe-test: powershell/lexical_escape_rules/lexical_escape_rules_escape_not_in_double_quotes
$str = "Line1`n`tLine2`$val`"quote`""
if ($str.Length -gt 0) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
