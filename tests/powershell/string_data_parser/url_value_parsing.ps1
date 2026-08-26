# vybe-test: powershell/string_data_parser/url_value_parsing
$str = "endpoint = https://api.service.io/v1?token=xyz&limit=50"
$ht = ConvertFrom-StringData -StringData $str
if ($ht["endpoint"] -ne "https://api.service.io/v1?token=xyz&limit=50") {
    Write-Host "FAIL: URL value parsing failed"
    exit 1
}
Write-Host "PASS"
exit 0
