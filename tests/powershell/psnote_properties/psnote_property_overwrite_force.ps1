# vybe-test: powershell/psnote_properties/psnote_property_overwrite_force
$obj = [pscustomobject]@{ Setting = "Old" }
$obj | Add-Member -NotePropertyName "Setting" -NotePropertyValue "New" -Force
if ($obj.Setting -ne "New") {
    Write-Host "FAIL: Add-Member -Force expected setting mutated to 'New'"
    exit 1
}
Write-Host "PASS"
exit 0
