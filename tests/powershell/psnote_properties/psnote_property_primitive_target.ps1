# vybe-test: powershell/psnote_properties/psnote_property_primitive_target
$str = "Text"
$wrapped = $str | Add-Member -NotePropertyName "Tag" -NotePropertyValue "WrappedStr" -PassThru
if ($wrapped.Tag -ne "WrappedStr") {
    Write-Host "FAIL: Add-Member NoteProperty to string expected Tag='WrappedStr'"
    exit 1
}
Write-Host "PASS"
exit 0
