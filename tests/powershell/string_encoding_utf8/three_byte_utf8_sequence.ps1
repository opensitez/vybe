# vybe-test: powershell/string_encoding_utf8/three_byte_utf8_sequence
$enc = [System.Text.Encoding]::UTF8
$bytes = $enc.GetBytes("`u{20AC}") # Euro € is 0xE2 0x82 0xAC
if ($bytes.Length -ne 3 -or $bytes[0] -ne 0xE2 -or $bytes[1] -ne 0x82 -or $bytes[2] -ne 0xAC) {
    Write-Host "FAIL: Three-byte UTF-8 sequence mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
