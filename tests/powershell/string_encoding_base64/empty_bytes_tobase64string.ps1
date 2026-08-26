# vybe-test: powershell/string_encoding_base64/empty_bytes_tobase64string
[byte[]]$empty = @()
$b64 = [System.Convert]::ToBase64String($empty)
if ($b64 -ne "") {
    Write-Host "FAIL: Base64 empty bytes failed"
    exit 1
}
Write-Host "PASS"
exit 0
