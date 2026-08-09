# vybe-test: powershell/psnote_properties/psnote_property_add_basic
$obj = [pscustomobject]@{}
$obj | Add-Member -NotePropertyName "Status" -NotePropertyValue "Active"
if ($obj.Status -ne "Active") {
    Write-Host "FAIL: Add-Member NoteProperty expected Status='Active', got '$($obj.Status)'"
    exit 1
}
Write-Host "PASS"
exit 0
