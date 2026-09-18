# vybe-test: powershell/test_json_cmdlet/test_json_primitive_boolean_valid
# Primitive booleans 'true' and 'false' are valid JSON
$trueRes = 'true' | Test-Json
$falseRes = 'false' | Test-Json

if ($trueRes -ne $true -or $falseRes -ne $true) {
    Write-Host "FAIL: primitive booleans failed validation, got true=$trueRes, false=$falseRes"
    exit 1
}

Write-Host "PASS"
exit 0
