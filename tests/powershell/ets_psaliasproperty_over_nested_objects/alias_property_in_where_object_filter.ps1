# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_in_where_object_filter
$items = @(
    [pscustomobject]@{ TotalCost = 150 },
    [pscustomobject]@{ TotalCost = 50 }
)
foreach ($it in $items) {
    $it.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Cost", "TotalCost"))
}
$expensive = @($items | Where-Object { $_.Cost -gt 100 })
if ($expensive.Length -ne 1 -or $expensive[0].TotalCost -ne 150) {
    Write-Host "FAIL: PSAliasProperty in Where-Object filter failed"
    exit 1
}
Write-Host "PASS"
exit 0
