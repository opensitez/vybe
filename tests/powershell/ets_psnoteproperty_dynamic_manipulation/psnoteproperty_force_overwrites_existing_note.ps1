# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_force_overwrites_existing_note
$obj = [pscustomobject]@{ Key = "old" }
$obj | Add-Member -NotePropertyName "Key" -NotePropertyValue "new" -Force
if ($obj.Key -ne "new") {
    Write-Host "FAIL: Add-Member -Force overwrite failed"
    exit 1
}
Write-Host "PASS"
exit 0
