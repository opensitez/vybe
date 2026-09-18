# vybe-test: powershell/test_json_cmdlet/test_json_schema_additional_properties_false
$schema = '{"type":"object","properties":{"known":{"type":"string"}},"additionalProperties":false}'

$validObj   = '{"known": "valid"}' | Test-Json -Schema $schema
$invalidObj = '{"known": "valid", "extraProperty": 99}' | Test-Json -Schema $schema -ErrorAction SilentlyContinue

if ($validObj -ne $true) {
    Write-Host "FAIL: object with only allowed properties failed validation"
    exit 1
}

if ($invalidObj -ne $false) {
    Write-Host "FAIL: object with extra property unexpectedly passed additionalProperties:false"
    exit 1
}

Write-Host "PASS"
exit 0
