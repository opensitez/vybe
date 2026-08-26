# vybe-test: powershell/string_encoding_utf8/empty_string_getbytes
$bytes = [System.Text.Encoding]::UTF8.GetBytes("")
if ($bytes.Length -ne 0) {
    Write-Host "FAIL: Empty string GetBytes must be empty array"
    exit 1
}
Write-Host "PASS"
exit 0
