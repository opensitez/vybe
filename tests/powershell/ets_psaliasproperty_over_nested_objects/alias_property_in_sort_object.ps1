# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_in_sort_object
$items = @(
    [pscustomobject]@{ RawScore = 30 },
    [pscustomobject]@{ RawScore = 10 },
    [pscustomobject]@{ RawScore = 20 }
)
foreach ($it in $items) {
    $it.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Score", "RawScore"))
}
$sorted = @($items | Sort-Object -Property Score)
if ($sorted[0].RawScore -ne 10 -or $sorted[2].RawScore -ne 30) {
    Write-Host "FAIL: PSAliasProperty in Sort-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
