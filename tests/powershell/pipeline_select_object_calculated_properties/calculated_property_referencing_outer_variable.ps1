# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_referencing_outer_variable
$rate = 1.25
$item = [pscustomobject]@{ Base = 100 }
$res = $item | Select-Object @{ N = "Converted"; E = { $_.Base * $rate } }
if ($res.Converted -ne 125.0) {
    Write-Host "FAIL: Calculated property outer variable closure failed"
    exit 1
}
Write-Host "PASS"
exit 0
