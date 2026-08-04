# vybe-test: powershell/strings/here_string_double
$name = "World"
$text = @"
Hello $name
Goodbye $name
"@
$result = $text -match "Hello World"
if ($result -ne $true) {
    Write-Host "FAIL: expected interpolation to work"
    exit 1
}
Write-Host "PASS"
exit 0
