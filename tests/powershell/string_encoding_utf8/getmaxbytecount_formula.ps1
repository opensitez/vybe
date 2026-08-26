# vybe-test: powershell/string_encoding_utf8/getmaxbytecount_formula
$enc = [System.Text.Encoding]::UTF8
$maxBytes = $enc.GetMaxByteCount(10)
if ($maxBytes -lt 10) {
    Write-Host "FAIL: GetMaxByteCount(10) must be at least 10"
    exit 1
}
Write-Host "PASS"
exit 0
