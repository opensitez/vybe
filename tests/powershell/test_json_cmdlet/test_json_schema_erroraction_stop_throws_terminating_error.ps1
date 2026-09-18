# vybe-test: powershell/test_json_cmdlet/test_json_schema_erroraction_stop_throws_terminating_error
$schema = '{"type":"object","properties":{"mandatory":{"type":"string"}},"required":["mandatory"]}'

# With -ErrorAction Stop, a schema validation failure converts into a terminating error
$threwExpected = $false
$caughtException = $null

try {
    '{"otherField": 123}' | Test-Json -Schema $schema -ErrorAction Stop
} catch {
    $threwExpected = $true
    $caughtException = $_
}

if (-not $threwExpected) {
    Write-Host "FAIL: Test-Json with -ErrorAction Stop did not throw on schema failure"
    exit 1
}

if ($caughtException.FullyQualifiedErrorId -notmatch "InvalidJsonAgainstSchemaDetailed|TestJsonCommand") {
    Write-Host "FAIL: unexpected error ID: $($caughtException.FullyQualifiedErrorId)"
    exit 1
}

Write-Host "PASS"
exit 0
