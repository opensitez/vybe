# vybe-test: powershell/variable_drives/variable_drive_local_prefix
$x = 10
function Local-DriveTest {
    $x = 20
    $local:x = 30
    return $local:x
}
$res = Local-DriveTest
if ($res -ne 30 -or $x -ne 10) {
    Write-Host "FAIL: \$local: drive prefix isolation failed"
    exit 1
}
Write-Host "PASS"
exit 0
