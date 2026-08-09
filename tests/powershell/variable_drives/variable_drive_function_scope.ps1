# vybe-test: powershell/variable_drives/variable_drive_function_scope
function Set-ScriptDriveVar {
    $script:driven = 77
}
Set-ScriptDriveVar
if ($driven -ne 77) {
    Write-Host "FAIL: \$script:driven set inside function expected 77, got $driven"
    exit 1
}
Write-Host "PASS"
exit 0
