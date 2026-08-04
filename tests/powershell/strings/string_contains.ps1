# vybe-test: powershell/strings/string_contains
$str = "PowerShell is great"
$result = $str.Contains("Shell")
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
