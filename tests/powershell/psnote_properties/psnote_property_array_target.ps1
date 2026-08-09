# vybe-test: powershell/psnote_properties/psnote_property_array_target
$arr = @(10, 20)
$arr | Add-Member -NotePropertyName "Meta" -NotePropertyValue "ArrayNote"
if ($arr.Meta -ne "ArrayNote") {
    Write-Host "FAIL: Add-Member NoteProperty to array target expected 'ArrayNote'"
    exit 1
}
Write-Host "PASS"
exit 0
