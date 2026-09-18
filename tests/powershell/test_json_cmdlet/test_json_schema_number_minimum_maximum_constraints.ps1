# vybe-test: powershell/test_json_cmdlet/test_json_schema_number_minimum_maximum_constraints
$schema = '{"type":"number","minimum":10,"maximum":50}'

$validVal = '25' | Test-Json -Schema $schema
$tooLow   = '5'  | Test-Json -Schema $schema -ErrorAction SilentlyContinue
$tooHigh  = '99' | Test-Json -Schema $schema -ErrorAction SilentlyContinue

if ($validVal -ne $true) {
    Write-Host "FAIL: in-bounds number failed validation"
    exit 1
}

if ($tooLow -ne $false -or $tooHigh -ne $false) {
    Write-Host "FAIL: out-of-bounds numbers did not return `$false, got tooLow=$tooLow, tooHigh=$tooHigh"
    exit 1
}

Write-Host "PASS"
exit 0
