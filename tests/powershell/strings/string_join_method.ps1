# vybe-test: powershell/strings/string_join_method
$parts = @("one", "two", "three")
$result = [string]::Join("-", $parts)
if ($result -ne "one-two-three") {
    Write-Host "FAIL: expected 'one-two-three', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
