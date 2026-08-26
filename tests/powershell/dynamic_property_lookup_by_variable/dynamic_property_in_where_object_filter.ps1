# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_in_where_object_filter
$filterField = "Active"
$items = @(
    [pscustomobject]@{ Active = $true; Name = "A" },
    [pscustomobject]@{ Active = $false; Name = "B" }
)
$activeItems = @($items | Where-Object { $_.$filterField })
if ($activeItems.Length -ne 1 -or $activeItems[0].Name -ne "A") {
    Write-Host "FAIL: Dynamic property in Where-Object filter failed"
    exit 1
}
Write-Host "PASS"
exit 0
