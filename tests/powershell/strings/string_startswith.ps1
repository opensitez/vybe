# vybe-test: powershell/strings/string_startswith
$str = "PowerShell"
$result = $str.StartsWith("Power")
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
