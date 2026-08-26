# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_overriding_existing_property_name
$item = [pscustomobject]@{ Status = "pending" }
$res = $item | Select-Object @{ N = "Status"; E = { $_.Status.ToUpper() } }
if ($res.Status -ne "PENDING") {
    Write-Host "FAIL: Calculated property overriding existing property name failed"
    exit 1
}
Write-Host "PASS"
exit 0
