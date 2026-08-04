# vybe-test: powershell/comparison/string_equal_case_sensitive
$result = ("hello" -ceq "HELLO")
if ($result -ne $false) {
    Write-Host "FAIL: expected False, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
