# vybe-test: powershell/strings/string_endswith
$str = "PowerShell"
$result = $str.EndsWith("Shell")
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
