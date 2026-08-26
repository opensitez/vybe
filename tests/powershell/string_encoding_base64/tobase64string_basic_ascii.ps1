# vybe-test: powershell/string_encoding_base64/tobase64string_basic_ascii
$bytes = [System.Text.Encoding]::UTF8.GetBytes("Hello")
$b64 = [System.Convert]::ToBase64String($bytes)
if ($b64 -ne "SGVsbG8=") {
    Write-Host "FAIL: ToBase64String failed, expected 'SGVsbG8=', got '$b64'"
    exit 1
}
Write-Host "PASS"
exit 0
