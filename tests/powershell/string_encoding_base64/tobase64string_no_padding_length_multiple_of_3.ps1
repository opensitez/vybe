# vybe-test: powershell/string_encoding_base64/tobase64string_no_padding_length_multiple_of_3
$bytes = [System.Text.Encoding]::UTF8.GetBytes("Man") # 3 bytes -> 4 chars, no padding
$b64 = [System.Convert]::ToBase64String($bytes)
if ($b64 -ne "TWFu") {
    Write-Host "FAIL: Base64 no padding failed, got '$b64'"
    exit 1
}
Write-Host "PASS"
exit 0
