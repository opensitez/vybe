# vybe-test: powershell/json_date_time_formats/rfc1123_http_date_format_in_json
$rfc = "Wed, 26 Aug 2026 12:00:00 GMT"
$json = @{ HttpDate = $rfc } | ConvertTo-Json
$obj = $json | ConvertFrom-Json
$dt = [datetime]::Parse($obj.HttpDate)
if ($dt.Year -ne 2026) {
    Write-Host "FAIL: RFC 1123 HTTP date in JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
