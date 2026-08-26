# vybe-test: powershell/string_data_parser/empty_value_allowed
$str = "emptyKey = "
$ht = ConvertFrom-StringData -StringData $str
if ($ht.ContainsKey("emptyKey") -ne $true -or $ht["emptyKey"] -ne "") {
    Write-Host "FAIL: Empty value parsing failed"
    exit 1
}
Write-Host "PASS"
exit 0
