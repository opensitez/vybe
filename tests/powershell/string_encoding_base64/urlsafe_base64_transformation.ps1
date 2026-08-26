# vybe-test: powershell/string_encoding_base64/urlsafe_base64_transformation
$orig = [System.Convert]::ToBase64String(@(0xFB, 0xFF, 0xFE)) # contains + and /
$urlSafe = $orig.Replace('+', '-').Replace('/', '_').TrimEnd('=')
$restored = $urlSafe.Replace('-', '+').Replace('_', '/')
while ($restored.Length % 4 -ne 0) { $restored += "=" }
$bytes = [System.Convert]::FromBase64String($restored)
if ($bytes[0] -ne 0xFB -or $bytes[1] -ne 0xFF -or $bytes[2] -ne 0xFE) {
    Write-Host "FAIL: URL-safe base64 roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
