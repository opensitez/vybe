# vybe-test: powershell/test_json_cmdlet/test_json_schema_required_property_missing_returns_false
$schema = '{"type":"object","properties":{"id":{"type":"integer"}},"required":["id"]}'

# JSON missing the required property fails schema validation and returns $false
$res = '{"name": "MissingId"}' | Test-Json -Schema $schema -ErrorAction SilentlyContinue

if ($res -ne $false) {
    Write-Host "FAIL: expected `$false when required property is missing, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
