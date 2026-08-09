# vybe-test: powershell/variable_drives/variable_drive_remove_item
$toRemove = "Removable"
Remove-Item "variable:toRemove"
if (Test-Path "variable:toRemove") {
    Write-Host "FAIL: Test-Path variable:toRemove expected false after Remove-Item"
    exit 1
}
Write-Host "PASS"
exit 0
