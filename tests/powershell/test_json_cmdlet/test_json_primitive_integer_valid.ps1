# vybe-test: powershell/test_json_cmdlet/test_json_primitive_integer_valid
# In RFC 8259, primitive literals (numbers) are valid top-level JSON documents
$res = '1048576' | Test-Json

if ($res -ne $true) {
    Write-Host "FAIL: primitive integer expected to be valid JSON, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
