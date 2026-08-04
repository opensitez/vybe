# vybe-test: powershell/strings/string_chars
$str = "abc"
$result = $str[0]
if ($result -ne 'a') {
    Write-Host "FAIL: expected 'a', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
