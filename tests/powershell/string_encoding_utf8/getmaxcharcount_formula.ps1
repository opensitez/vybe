# vybe-test: powershell/string_encoding_utf8/getmaxcharcount_formula
$enc = [System.Text.Encoding]::UTF8
$maxChars = $enc.GetMaxCharCount(10)
if ($maxChars -lt 1) {
    Write-Host "FAIL: GetMaxCharCount failed"
    exit 1
}
Write-Host "PASS"
exit 0
