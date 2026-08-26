# vybe-test: powershell/string_encoding_base64/frombase64string_roundtrip
$originalText = "PowerShell Base64 Test String"
$bytes = [System.Text.Encoding]::UTF8.GetBytes($originalText)
$b64 = [System.Convert]::ToBase64String($bytes)
$decodedBytes = [System.Convert]::FromBase64String($b64)
$decodedText = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
if ($decodedText -ne $originalText) {
    Write-Host "FAIL: FromBase64String roundtrip failed, got '$decodedText'"
    exit 1
}
Write-Host "PASS"
exit 0
