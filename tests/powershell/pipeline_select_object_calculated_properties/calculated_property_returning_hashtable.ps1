# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_returning_hashtable
$item = [pscustomobject]@{ Key = "K"; Val = "V" }
$res = $item | Select-Object @{ N = "Dict"; E = { @{ ($_.Key) = $_.Val } } }
if ($res.Dict["K"] -ne "V") {
    Write-Host "FAIL: Calculated property returning hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
