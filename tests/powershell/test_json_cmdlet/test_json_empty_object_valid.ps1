# vybe-test: powershell/test_json_cmdlet/test_json_empty_object_valid
# An empty JSON object '{}' is valid JSON
$res = '{}' | Test-Json

if ($res -ne $true) {
    Write-Host "FAIL: empty object '{}' expected to be valid JSON, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
