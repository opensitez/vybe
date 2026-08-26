# vybe-test: powershell/string_encoding_base64/tobase64string_two_pad_chars
$bytes = [System.Text.Encoding]::UTF8.GetBytes("M") # 1 byte -> 2 '=' pad
$b64 = [System.Convert]::ToBase64String($bytes)
if ($b64 -ne "TQ==") {
    Write-Host "FAIL: Base64 double pad failed, got '$b64'"
    exit 1
}
Write-Host "PASS"
exit 0
