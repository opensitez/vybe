# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_pipeline_selection
$targetField = "Price"
$items = @(
    [pscustomobject]@{ Name = "Item1"; Price = 10 },
    [pscustomobject]@{ Name = "Item2"; Price = 20 }
)
$prices = @($items | ForEach-Object { $_.$targetField })
if ($prices.Length -ne 2 -or $prices[0] -ne 10 -or $prices[1] -ne 20) {
    Write-Host "FAIL: Dynamic property pipeline selection failed"
    exit 1
}
Write-Host "PASS"
exit 0
