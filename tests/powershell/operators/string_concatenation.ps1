# vybe-test: powershell/operators/string_concatenation
$result = "Hello" + " " + "World"
if ($result -ne "Hello World") {
    Write-Host "FAIL: expected 'Hello World', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
