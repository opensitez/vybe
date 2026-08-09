# vybe-test: powershell/variable_drives/variable_drive_env_set
$env:VYBE_TEST_MUTATE = "Original"
$env:VYBE_TEST_MUTATE = "Updated"
if ($env:VYBE_TEST_MUTATE -ne "Updated") {
    Write-Host "FAIL: environment variable mutation failed"
    exit 1
}
Write-Host "PASS"
exit 0
