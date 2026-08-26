# vybe-test: powershell/string_data_parser/unicode_characters_in_key_and_value
$str = "caf`u{00E9} = resu`u{00E9}m"
$ht = ConvertFrom-StringData -StringData $str
if ($ht["caf`u{00E9}"] -ne "resu`u{00E9}m") {
    Write-Host "FAIL: Unicode in key/value failed"
    exit 1
}
Write-Host "PASS"
exit 0
