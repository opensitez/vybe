# vybe-test: powershell/string_encoding_utf8/roundtrip_multilingual_string
$orig = "Hello 世界 🌍"
$bytes = [System.Text.Encoding]::UTF8.GetBytes($orig)
$reconstructed = [System.Text.Encoding]::UTF8.GetString($bytes)
if ($orig -ne $reconstructed) {
    Write-Host "FAIL: Multilingual UTF-8 roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
