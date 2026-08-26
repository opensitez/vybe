# vybe-test: powershell/string_encoding_hex_conversions/roundtrip_text_via_utf8_hex
$orig = "PowerShell"
$bytes = [System.Text.Encoding]::UTF8.GetBytes($orig)
$hex = [System.Convert]::ToHexString($bytes)
$recoveredBytes = [System.Convert]::FromHexString($hex)
$recovered = [System.Text.Encoding]::UTF8.GetString($recoveredBytes)
if ($orig -ne $recovered) {
    Write-Host "FAIL: Text to Hex to Text roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
