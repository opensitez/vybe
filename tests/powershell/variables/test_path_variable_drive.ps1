# vybe-test: powershell/variables/test_path_variable_drive
# Test-Path on variable: provider drive accurately reports whether a variable exists
$activeVar = 42

$exists = Test-Path "variable:activeVar"
$missing = Test-Path "variable:definitely_nonexistent_variable_xyz"

if (-not $exists) {
    Write-Host "FAIL: Test-Path variable:activeVar returned false"
    exit 1
}

if ($missing) {
    Write-Host "FAIL: Test-Path variable:definitely_nonexistent_variable_xyz returned true"
    exit 1
}

Write-Host "PASS"
exit 0
