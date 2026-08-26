# vybe-test: powershell/string_data_parser/colon_separator_not_treated_as_equals
$str = "key:with:colons = val"
$ht = ConvertFrom-StringData -StringData $str
if ($ht["key:with:colons"] -ne "val") {
    Write-Host "FAIL: Key with colons failed"
    exit 1
}
Write-Host "PASS"
exit 0
