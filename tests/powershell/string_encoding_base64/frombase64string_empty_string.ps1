# vybe-test: powershell/string_encoding_base64/frombase64string_empty_string
$bytes = [System.Convert]::FromBase64String("")
if ($bytes.Length -ne 0) {
    Write-Host "FAIL: FromBase64String empty string failed"
    exit 1
}
Write-Host "PASS"
exit 0
