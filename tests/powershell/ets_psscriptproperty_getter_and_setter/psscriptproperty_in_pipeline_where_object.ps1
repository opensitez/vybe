# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_in_pipeline_where_object
$items = @(
    [pscustomobject]@{ Qty = 5; UnitPrice = 10 },
    [pscustomobject]@{ Qty = 2; UnitPrice = 5 }
)
foreach ($it in $items) {
    $it.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Total", { $this.Qty * $this.UnitPrice }))
}
$expensive = @($items | Where-Object { $_.Total -ge 50 })
if ($expensive.Length -ne 1 -or $expensive[0].Total -ne 50) {
    Write-Host "FAIL: PSScriptProperty in Where-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
