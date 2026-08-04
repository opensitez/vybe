# vybe-test: powershell/strings/string_insert
$str = "HelloWorld"
$result = $str.Insert(5, " ")
if ($result -ne "Hello World") {
    Write-Host "FAIL: expected 'Hello World', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
