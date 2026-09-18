# vybe-test: powershell/test_json_cmdlet/test_json_valid_object_returns_true
# Test-Json validates that a string contains well-formed JSON and returns $true
$res = '{"name": "Alice", "age": 30}' | Test-Json

if ($res -ne $true) {
    Write-Host "FAIL: expected `$true for valid JSON object, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
