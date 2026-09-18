# vybe-test: powershell/test_json_cmdlet/test_json_schema_array_min_items_constraint
$schema = '{"type":"array","minItems":3}'

$passArray = '[1, 2, 3]' | Test-Json -Schema $schema
$failArray = '[1, 2]'    | Test-Json -Schema $schema -ErrorAction SilentlyContinue

if ($passArray -ne $true) {
    Write-Host "FAIL: array with 3 items failed minItems:3 constraint"
    exit 1
}

if ($failArray -ne $false) {
    Write-Host "FAIL: array with 2 items unexpectedly passed minItems:3 constraint"
    exit 1
}

Write-Host "PASS"
exit 0
