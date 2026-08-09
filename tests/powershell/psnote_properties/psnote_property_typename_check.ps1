# vybe-test: powershell/psnote_properties/psnote_property_typename_check
$obj = [pscustomobject]@{}
$obj | Add-Member -NotePropertyName "Prop" -NotePropertyValue "Val" -TypeName "CustomTypeName"
if ($obj.psobject.TypeNames[0] -ne "CustomTypeName") {
    Write-Host "FAIL: Add-Member -TypeName expected CustomTypeName, got $($obj.psobject.TypeNames[0])"
    exit 1
}
Write-Host "PASS"
exit 0
