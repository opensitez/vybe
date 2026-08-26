# vybe-test: powershell/string_data_parser/trailing_spaces_in_multiline_data
$str = "k1 = v1   `nk2 = v2   "
$ht = ConvertFrom-StringData -StringData $str
if ($ht["k1"] -ne "v1" -or $ht["k2"] -ne "v2") {
    Write-Host "FAIL: Trailing spaces in multiline data failed"
    exit 1
}
Write-Host "PASS"
exit 0
