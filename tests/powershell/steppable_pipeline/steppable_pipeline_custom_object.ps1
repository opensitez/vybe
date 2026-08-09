# vybe-test: powershell/steppable_pipeline/steppable_pipeline_custom_object
$sb = { process { $_.Tag } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process([pscustomobject]@{ Tag = "SpObj" })
$sp.End()
if ($res -ne "SpObj") {
    Write-Host "FAIL: SteppablePipeline custom object expected SpObj, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
