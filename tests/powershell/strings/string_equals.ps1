# vybe-test: powershell/strings/string_equals
$result = "hello".Equals("hello")
if ($result -ne $true) {
    Write-Host "FAIL: expected True for string equality"
    exit 1
}
$result2 = "hello".Equals("HELLO")
if ($result2 -ne $false) {
    Write-Host "FAIL: expected False for case-insensitive equality"
    exit 1
}
Write-Host "PASS"
exit 0
