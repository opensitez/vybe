# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_l_key_alias
$items = @([pscustomobject]@{ Code = "prod" })
$res = $items | Select-Object @{ l = "Upper"; e = { $_.Code.ToUpper() } }
if ($res.Upper -ne "PROD") {
    Write-Host "FAIL: Calculated property l key alias failed"
    exit 1
}
Write-Host "PASS"
exit 0
