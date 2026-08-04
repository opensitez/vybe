# vybe-test: powershell/comparison/like_operator
$result = "PowerShell" -like "Power*"
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
