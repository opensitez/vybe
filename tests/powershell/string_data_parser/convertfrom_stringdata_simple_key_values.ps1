# vybe-test: powershell/string_data_parser/convertfrom_stringdata_simple_key_values
$str = @"
key1 = value1
key2 = value2
"@
$ht = ConvertFrom-StringData -StringData $str
if ($ht["key1"] -ne "value1" -or $ht["key2"] -ne "value2" -or $ht.Count -ne 2) {
    Write-Host "FAIL: ConvertFrom-StringData simple parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
