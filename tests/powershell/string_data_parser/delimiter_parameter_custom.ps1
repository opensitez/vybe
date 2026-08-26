# vybe-test: powershell/string_data_parser/delimiter_parameter_custom
$str = "key:value"
$ht = ConvertFrom-StringData -StringData $str -Delimiter ":"
if ($ht["key"] -ne "value") {
    Write-Host "FAIL: Custom delimiter ':' failed"
    exit 1
}
Write-Host "PASS"
exit 0
