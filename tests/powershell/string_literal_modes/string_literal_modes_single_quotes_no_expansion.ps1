# vybe-test: powershell/string_literal_modes/string_literal_modes_single_quotes_no_expansion
$str = "Line1`n`tLine2`$val`"quote`""
if ($str.Length -gt 0) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
