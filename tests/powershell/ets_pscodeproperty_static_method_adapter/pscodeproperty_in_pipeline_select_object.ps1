# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_in_pipeline_select_object
class PipeCode {
    static [int]GetSquare([psobject]$i) { return $i.Num * $i.Num }
}
$items = @([pscustomobject]@{ Num = 5 }, [pscustomobject]@{ Num = 6 })
$m = [PipeCode].GetMethod("GetSquare")
foreach ($it in $items) {
    $it.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("Square", $m))
}
$res = @($items | Select-Object -ExpandProperty Square)
if ($res[0] -ne 25 -or $res[1] -ne 36) {
    Write-Host "FAIL: PSCodeProperty in Select-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
