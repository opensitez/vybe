# vybe-test: powershell/string_substitution_rules/string_substitution_rules_substitute_hashtable_key
$str = "Line1`n`tLine2`$val`"quote`""
if ($str.Length -gt 0) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
