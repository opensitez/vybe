# vybe-test: powershell/string_encoding_utf8/getbytes_ascii_characters
$enc = [System.Text.Encoding]::UTF8
$bytes = $enc.GetBytes("ABC")
if ($bytes.Length -ne 3 -or $bytes[0] -ne 65 -or $bytes[1] -ne 66 -or $bytes[2] -ne 67) {
    Write-Host "FAIL: UTF8 GetBytes ASCII mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
