# vybe-test: powershell/strings/string_lastindexof
$str = "hello world hello"
$index = $str.LastIndexOf("hello")
if ($index -ne 12) {
    Write-Host "FAIL: expected 12, got $index"
    exit 1
}
Write-Host "PASS"
exit 0
