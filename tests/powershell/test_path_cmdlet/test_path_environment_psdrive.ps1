# vybe-test: powershell/test_path_cmdlet/test_path_environment_psdrive
# Test-Path checks existence of environment variables via the Env: PSDrive
$exists = Test-Path "Env:PATH"
$missing = Test-Path "Env:COMPLETELY_NONEXISTENT_ENV_VAR_XYZ_99"

if ($exists -ne $true) {
    Write-Host "FAIL: Test-Path 'Env:PATH' returned `$false"
    exit 1
}

if ($missing -ne $false) {
    Write-Host "FAIL: Test-Path on missing Env variable returned `$true"
    exit 1
}

Write-Host "PASS"
exit 0
