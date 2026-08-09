# vybe-test: powershell/psnote_properties/psnote_property_null_value
$obj = [pscustomobject]@{}
$obj | Add-Member -NotePropertyName "NullProp" -NotePropertyValue $null
$prop = $obj.psobject.Properties["NullProp"]
if ($prop -eq $null -or $prop.Value -ne $null) {
    Write-Host "FAIL: NoteProperty with null value registration failed"
    exit 1
}
Write-Host "PASS"
exit 0
