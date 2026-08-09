# vybe-test: powershell/psnote_properties/psnote_property_remove
$obj = [pscustomobject]@{ DynamicProp = "Temp" }
$obj.psobject.Properties.Remove("DynamicProp")
if ($obj.psobject.Properties["DynamicProp"] -ne $null) {
    Write-Host "FAIL: PSNoteProperty removal failed"
    exit 1
}
Write-Host "PASS"
exit 0
