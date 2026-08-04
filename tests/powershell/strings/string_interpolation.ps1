# vybe-test: powershell/strings/string_interpolation
$name = "World"
$result = "Hello, $name!"
if ($result -ne "Hello, World!") {
    Write-Host "FAIL: expected 'Hello, World!', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
