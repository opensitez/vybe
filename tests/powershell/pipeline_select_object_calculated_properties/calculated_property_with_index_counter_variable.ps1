# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_with_index_counter_variable
$items = @("A", "B", "C")
$res = @($items | Select-Object @{ N = "Upper"; E = { $_.ToUpper() } })
if ($res.Length -ne 3 -or $res[0].Upper -ne "A" -or $res[2].Upper -ne "C") {
    Write-Host "FAIL: Calculated property failed"
    exit 1
}
Write-Host "PASS"
exit 0
