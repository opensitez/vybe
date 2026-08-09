# vybe-test: powershell/steppable_pipeline/steppable_pipeline_cmdletbinding
$sb = {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$InputVal)
    process { $InputVal * 100 }
}
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process(7)
$sp.End()
if ($res -ne 700) {
    Write-Host "FAIL: SteppablePipeline CmdletBinding expected 700, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
