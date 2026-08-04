# vybe-test: powershell/strings/string_replace
$str = "Hello World"
$result = $str.Replace("World", "PowerShell")
if ($result -ne "Hello PowerShell") {
    Write-Host "FAIL: expected 'Hello PowerShell', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
