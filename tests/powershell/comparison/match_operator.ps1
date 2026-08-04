# vybe-test: powershell/comparison/match_operator
$result = "hello123" -match "\d+"
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
