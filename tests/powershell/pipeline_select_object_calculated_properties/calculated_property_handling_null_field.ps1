# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_handling_null_field
$item = [pscustomobject]@{ Name = $null }
$res = $item | Select-Object @{ N = "SafeName"; E = { if ($null -eq $_.Name) { "Anonymous" } else { $_.Name } } }
if ($res.SafeName -ne "Anonymous") {
    Write-Host "FAIL: Calculated property null field handling failed"
    exit 1
}
Write-Host "PASS"
exit 0
