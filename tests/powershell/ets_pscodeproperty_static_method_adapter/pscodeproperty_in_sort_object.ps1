# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_in_sort_object
class SortAdapter {
    static [int]GetNegScore([psobject]$i) { return -$i.Score }
}
$items = @(
    [pscustomobject]@{ Score = 10 },
    [pscustomobject]@{ Score = 30 },
    [pscustomobject]@{ Score = 20 }
)
$m = [SortAdapter].GetMethod("GetNegScore")
foreach ($it in $items) {
    $it.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("NegScore", $m))
}
$sorted = @($items | Sort-Object -Property NegScore)
if ($sorted[0].Score -ne 30 -or $sorted[2].Score -ne 10) {
    Write-Host "FAIL: PSCodeProperty in Sort-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
