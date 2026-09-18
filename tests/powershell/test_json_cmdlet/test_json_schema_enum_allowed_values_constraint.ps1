# vybe-test: powershell/test_json_cmdlet/test_json_schema_enum_allowed_values_constraint
$schema = '{"type":"string","enum":["North","South","East","West"]}'

$passVal = '"East"' | Test-Json -Schema $schema
$failVal = '"Up"'   | Test-Json -Schema $schema -ErrorAction SilentlyContinue

if ($passVal -ne $true) {
    Write-Host "FAIL: allowed enum value failed validation"
    exit 1
}

if ($failVal -ne $false) {
    Write-Host "FAIL: disallowed enum value unexpectedly passed validation"
    exit 1
}

Write-Host "PASS"
exit 0
