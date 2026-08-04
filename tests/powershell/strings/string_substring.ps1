# vybe-test: powershell/strings/string_substring
$str = "PowerShell"
$result = $str.Substring(0, 5)
if ($result -ne "Power") {
    Write-Host "FAIL: expected 'Power', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
