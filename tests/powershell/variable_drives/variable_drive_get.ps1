# vybe-test: powershell/variable_drives/variable_drive_get
$myVar = "DriveTest"
$val = $variable:myVar
if ($val -ne "DriveTest") {
    Write-Host "FAIL: \$variable:myVar access expected 'DriveTest', got '$val'"
    exit 1
}
Write-Host "PASS"
exit 0
