# vybe-test: powershell/strings/string_remove
$str = "Hello World"
$result = $str.Remove(5, 6)  # Remove " World"
if ($result -ne "Hello") {
    Write-Host "FAIL: expected 'Hello', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
