# vybe-test: powershell/test_json_cmdlet/test_json_schema_type_mismatch_returns_false
$schema = '{"type":"object","properties":{"count":{"type":"integer"}}}'

# Supplying a string where integer is required fails schema validation
$res = '{"count": "not_a_number"}' | Test-Json -Schema $schema -ErrorAction SilentlyContinue

if ($res -ne $false) {
    Write-Host "FAIL: expected `$false for type mismatch against schema, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
