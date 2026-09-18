# vybe-test: powershell/test_json_cmdlet/test_json_schema_regex_pattern_constraint
# Pattern regex checks format: 3 uppercase letters followed by 3 digits (e.g. ABC123)
$schema = '{"type":"string","pattern":"^[A-Z]{3}[0-9]{3}$"}'

$validFormat   = '"XYZ789"' | Test-Json -Schema $schema
$invalidFormat = '"xyz789"' | Test-Json -Schema $schema -ErrorAction SilentlyContinue

if ($validFormat -ne $true) {
    Write-Host "FAIL: matching pattern failed validation"
    exit 1
}

if ($invalidFormat -ne $false) {
    Write-Host "FAIL: non-matching pattern unexpectedly passed validation"
    exit 1
}

Write-Host "PASS"
exit 0
