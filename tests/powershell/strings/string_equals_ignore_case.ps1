# vybe-test: powershell/strings/string_equals_ignore_case
$result = "hello".Equals("HELLO", "OrdinalIgnoreCase")
if ($result -ne $true) {
    Write-Host "FAIL: expected True for case-insensitive equality"
    exit 1
}
Write-Host "PASS"
exit 0
