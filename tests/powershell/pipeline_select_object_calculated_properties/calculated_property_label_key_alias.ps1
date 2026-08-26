# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_label_key_alias
$obj = [pscustomobject]@{ Raw = "hello" }
$res = $obj | Select-Object @{ l = "Upper"; e = { $_.Raw.ToUpper() } }
if ($res.Upper -ne "HELLO") {
    Write-Host "FAIL: Calculated property Label key alias failed"
    exit 1
}
Write-Host "PASS"
exit 0
