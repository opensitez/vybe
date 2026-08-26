# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_dynamic_invocations_in_loop
class LoopCode { static [int]GetInc([psobject]$i) { return $i.Val + 1 } }
$m = [LoopCode].GetMethod("GetInc")
$sum = 0
for ($k = 0; $k -lt 5; $k++) {
    $o = [pscustomobject]@{ Val = $k }
    $o.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("Inc", $m))
    $sum += $o.Inc
} # (1 + 2 + 3 + 4 + 5) = 15
if ($sum -ne 15) {
    Write-Host "FAIL: PSCodeProperty in loop failed, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
