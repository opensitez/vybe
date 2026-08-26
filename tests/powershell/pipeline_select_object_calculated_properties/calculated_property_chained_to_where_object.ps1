# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_chained_to_where_object
$items = @(
    [pscustomobject]@{ Price = 10 },
    [pscustomobject]@{ Price = 25 }
)
$res = @($items | Select-Object @{ N = "TaxPrice"; E = { $_.Price * 1.2 } } | Where-Object { $_.TaxPrice -gt 20 })
if ($res.Length -ne 1 -or $res[0].TaxPrice -ne 30.0) {
    Write-Host "FAIL: Calculated property chained to Where-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
