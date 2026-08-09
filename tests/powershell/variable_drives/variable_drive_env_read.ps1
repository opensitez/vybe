# vybe-test: powershell/variable_drives/variable_drive_env_read
$env:VYBE_TEST_ENV_VAR = "Active"
if ($env:VYBE_TEST_ENV_VAR -ne "Active") {
    Write-Host "FAIL: read \$env:VYBE_TEST_ENV_VAR expected 'Active', got '$env:VYBE_TEST_ENV_VAR'"
    exit 1
}
Write-Host "PASS"
exit 0
