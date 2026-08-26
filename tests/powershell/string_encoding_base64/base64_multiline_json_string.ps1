# vybe-test: powershell/string_encoding_base64/base64_multiline_json_string
$json = '{"name":"Alice","admin":true}'
$b64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($json))
$recovered = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($b64))
if ($recovered -ne $json) {
    Write-Host "FAIL: JSON base64 roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
