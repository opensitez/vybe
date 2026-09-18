# vybe-test: powershell/test_json_cmdlet/test_json_nested_object_schema_validation
$schema = @'
{
  "type": "object",
  "properties": {
    "account": {
      "type": "object",
      "properties": {
        "username": { "type": "string" },
        "tier": { "type": "integer" }
      },
      "required": ["username", "tier"]
    }
  },
  "required": ["account"]
}
'@

$validNested   = '{"account": {"username": "superuser", "tier": 1}}' | Test-Json -Schema $schema
$invalidNested = '{"account": {"username": "superuser"}}' | Test-Json -Schema $schema -ErrorAction SilentlyContinue

if ($validNested -ne $true) {
    Write-Host "FAIL: valid nested object failed schema validation"
    exit 1
}

if ($invalidNested -ne $false) {
    Write-Host "FAIL: invalid nested object (missing tier) unexpectedly passed validation"
    exit 1
}

Write-Host "PASS"
exit 0
