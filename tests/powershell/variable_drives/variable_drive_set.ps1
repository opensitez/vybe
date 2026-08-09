# vybe-test: powershell/variable_drives/variable_drive_set
$variable:newDriveVar = 12345
if ($newDriveVar -ne 12345) {
    Write-Host "FAIL: assignment to \$variable:newDriveVar failed, got $newDriveVar"
    exit 1
}
Write-Host "PASS"
exit 0
