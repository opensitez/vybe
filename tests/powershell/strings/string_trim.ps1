# vybe-test: powershell/strings/string_trim
$str = "  hello  "
$result = $str.Trim()
if ($result -ne "hello") {
    Write-Host "FAIL: expected 'hello', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
