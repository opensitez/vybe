# vybe-test: powershell/process_record_streaming/process_block_steppable_pipeline_step_by_step
# A steppable pipeline enables manual feed-by-feed execution of the pipeline processing loop
$scriptblock = { ForEach-Object { $_ * 3 } }

$steppable = $scriptblock.GetSteppablePipeline()
$steppable.Begin($true)

$out1 = $steppable.Process(10)
$out2 = $steppable.Process(20)
$out3 = $steppable.Process(30)

$steppable.End()

if ($out1 -ne 30) {
    Write-Host "FAIL: step 1 expected 30, got $out1"
    exit 1
}

if ($out2 -ne 60) {
    Write-Host "FAIL: step 2 expected 60, got $out2"
    exit 1
}

if ($out3 -ne 90) {
    Write-Host "FAIL: step 3 expected 90, got $out3"
    exit 1
}

Write-Host "PASS"
exit 0
