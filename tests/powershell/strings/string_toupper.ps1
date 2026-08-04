# vybe-test: powershell/strings/string_toupper
$str = "hello"
$result = $str.ToUpper()
if ($result -ne "HELLO") {
    Write-Host "FAIL: expected 'HELLO', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
