# vybe-test: powershell/comparison/string_equal_case_insensitive
$result = ("hello" -eq "HELLO")
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
