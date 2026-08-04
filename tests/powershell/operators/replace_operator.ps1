# vybe-test: powershell/operators/replace_operator
$text = "Hello World"
$result = $text -replace "World", "PowerShell"
if ($result -ne "Hello PowerShell") {
    Write-Host "FAIL: expected 'Hello PowerShell', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
