# vybe-test: powershell/string_encoding_base64/tobase64string_one_pad_char
$bytes = [System.Text.Encoding]::UTF8.GetBytes("Ma") # 2 bytes -> 1 '=' pad
$b64 = [System.Convert]::ToBase64String($bytes)
if ($b64 -ne "TWE=") {
    Write-Host "FAIL: Base64 single pad failed, got '$b64'"
    exit 1
}
Write-Host "PASS"
exit 0
