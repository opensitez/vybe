# vybe-test: powershell/test_json_cmdlet/test_json_empty_array_valid
# An empty JSON array '[]' is valid JSON
$res = '[]' | Test-Json

if ($res -ne $true) {
    Write-Host "FAIL: empty array '[]' expected to be valid JSON, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
