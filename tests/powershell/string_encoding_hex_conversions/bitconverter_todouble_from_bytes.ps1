# vybe-test: powershell/string_encoding_hex_conversions/bitconverter_todouble_from_bytes
$orig = 3.14159
$bytes = [System.BitConverter]::GetBytes($orig)
$recovered = [System.BitConverter]::ToDouble($bytes, 0)
if ($orig -ne $recovered) {
    Write-Host "FAIL: Double to bytes to double roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
