# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_with_guid_generation
$items = @(1, 2)
$res = @($items | Select-Object @{ N = "Id"; E = { [guid]::NewGuid() } })
if ($res[0].Id -isnot [guid] -or $res[0].Id -eq $res[1].Id) {
    Write-Host "FAIL: Calculated property GUID generation failed"
    exit 1
}
Write-Host "PASS"
exit 0
