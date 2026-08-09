# vybe-test: powershell/psnote_properties/psnote_property_hashtable_target
$h = @{ A = 1 }
$h | Add-Member -NotePropertyName "Attached" -NotePropertyValue "Ext"
if ($h.Attached -ne "Ext") {
    Write-Host "FAIL: Add-Member NoteProperty to hashtable target expected 'Ext'"
    exit 1
}
Write-Host "PASS"
exit 0
