# vybe-test: powershell/variable_drives/variable_drive_hashtable_item
$variable:map = @{ Key = "Value" }
if ($map.Key -ne "Value") {
    Write-Host "FAIL: \$variable: map expected Key=Value"
    exit 1
}
Write-Host "PASS"
exit 0
