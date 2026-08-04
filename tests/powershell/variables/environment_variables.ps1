# vybe-test: powershell/variables/environment_variables
$env:TEST_VAR = "test_value"
$result = $env:TEST_VAR
if ($result -ne "test_value") {
    Write-Host "FAIL: expected 'test_value', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
