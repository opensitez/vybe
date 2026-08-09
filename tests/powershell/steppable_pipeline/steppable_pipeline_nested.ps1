# vybe-test: powershell/steppable_pipeline/steppable_pipeline_nested
$innerSb = { process { $_ + 1 } }
$outerSb = {
    process {
        $spInner = $using:innerSb.GetSteppablePipeline()
        $spInner.Begin($true)
        $out = $spInner.Process($_)
        $spInner.End()
        return $out
    }
}.GetClosure()
$spOuter = $outerSb.GetSteppablePipeline()
$spOuter.Begin($true)
$res = $spOuter.Process(10)
$spOuter.End()
if ($res -ne 11) {
    Write-Host "FAIL: nested SteppablePipeline expected 11, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
