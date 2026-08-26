# vybe-test: powershell/string_encoding_utf8/two_byte_utf8_sequence
$enc = [System.Text.Encoding]::UTF8
$bytes = $enc.GetBytes("`u{00E9}") # é is 0xC3 0xA9
if ($bytes.Length -ne 2 -or $bytes[0] -ne 0xC3 -or $bytes[1] -ne 0xA9) {
    Write-Host "FAIL: Two-byte UTF-8 sequence mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
