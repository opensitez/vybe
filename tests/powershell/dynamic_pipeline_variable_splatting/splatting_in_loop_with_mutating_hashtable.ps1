# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_in_loop_with_mutating_hashtable
function Target-LoopSplat {
    param([int]$Step)
    return $Step * 10
}
$p = @{ Step = 0 }
$results = [System.Collections.Generic.List[int]]::new()
for ($i = 1; $i -le 3; $i++) {
    $p.Step = $i
    $results.Add((Target-LoopSplat @p))
}
if ($results[0] -ne 10 -or $results[1] -ne 20 -or $results[2] -ne 30) {
    Write-Host "FAIL: Splatting in loop with mutating hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
