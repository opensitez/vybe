# vybe-test: powershell/variable_drives/variable_drive_test_path
$pathVar = "Exists"
if (-not (Test-Path "variable:pathVar")) {
    Write-Host "FAIL: Test-Path variable:pathVar expected true"
    exit 1
}
if (Test-Path "variable:nonExistentVar999") {
    Write-Host "FAIL: Test-Path non-existent variable expected false"
    exit 1
}
Write-Host "PASS"
exit 0
