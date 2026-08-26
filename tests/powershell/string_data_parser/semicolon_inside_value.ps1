# vybe-test: powershell/string_data_parser/semicolon_inside_value
$str = "items = one;two;three"
$ht = ConvertFrom-StringData -StringData $str
if ($ht["items"] -ne "one;two;three") {
    Write-Host "FAIL: Semicolon inside value failed"
    exit 1
}
Write-Host "PASS"
exit 0
