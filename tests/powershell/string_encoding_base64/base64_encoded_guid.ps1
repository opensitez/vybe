# vybe-test: powershell/string_encoding_base64/base64_encoded_guid
$g = [guid]::NewGuid()
$b64 = [System.Convert]::ToBase64String($g.ToByteArray())
$bytes = [System.Convert]::FromBase64String($b64)
$reconstructed = [guid]::new($bytes)
if ($g -ne $reconstructed) {
    Write-Host "FAIL: Base64 GUID roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
