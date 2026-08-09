# vybe-test: powershell/variable_drives/variable_drive_child_items
$prefixTest = "Data1"
$vars = Get-ChildItem "variable:prefixTest*"
if ($vars.Count -lt 1 -or $vars[0].Value -ne "Data1") {
    Write-Host "FAIL: Get-ChildItem variable:prefixTest* failed"
    exit 1
}
Write-Host "PASS"
exit 0
