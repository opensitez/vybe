# vybe-test: powershell/test_json_cmdlet/test_json_direct_json_parameter
# Test-Json accepts the JSON string directly via the -Json parameter
$res = Test-Json -Json '{"directBinding": true, "code": 200}'

if ($res -ne $true) {
    Write-Host "FAIL: direct -Json parameter failed validation, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
