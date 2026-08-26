# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_with_array_parameter_in_hashtable
function Target-ArrParam {
    param([string[]]$Tags)
    return $Tags.Length
}
$p = @{ Tags = @("t1", "t2", "t3") }
$res = Target-ArrParam @p
if ($res -ne 3) {
    Write-Host "FAIL: Splatting array parameter inside hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
