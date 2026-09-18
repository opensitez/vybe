# vybe-test: powershell/test_json_cmdlet/test_json_schema_required_property_success
$schema = '{"type":"object","properties":{"id":{"type":"integer"}},"required":["id"]}'

# JSON containing the required property passes schema validation
$res = '{"id": 42}' | Test-Json -Schema $schema

if ($res -ne $true) {
    Write-Host "FAIL: schema validation with required property present failed, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
