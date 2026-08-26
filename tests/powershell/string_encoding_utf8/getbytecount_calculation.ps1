# vybe-test: powershell/string_encoding_utf8/getbytecount_calculation
$enc = [System.Text.Encoding]::UTF8
$count = $enc.GetByteCount("ABC€") # 1 + 1 + 1 + 3 = 6 bytes
if ($count -ne 6) {
    Write-Host "FAIL: GetByteCount expected 6, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
