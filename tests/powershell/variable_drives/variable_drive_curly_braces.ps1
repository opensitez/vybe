# vybe-test: powershell/variable_drives/variable_drive_curly_braces
${variable:var with spaces} = "SpaceVal"
if (${var with spaces} -ne "SpaceVal") {
    Write-Host "FAIL: curly brace variable drive assignment failed"
    exit 1
}
Write-Host "PASS"
exit 0
