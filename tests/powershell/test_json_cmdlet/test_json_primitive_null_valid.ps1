# vybe-test: powershell/test_json_cmdlet/test_json_primitive_null_valid
# Primitive literal 'null' is a valid JSON document
$res = 'null' | Test-Json

if ($res -ne $true) {
    Write-Host "FAIL: primitive 'null' expected to be valid JSON, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
