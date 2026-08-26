# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_streaming_evaluation_lazy_side_effect
$streamed = [System.Collections.Generic.List[int]]::new()
$sideStore = $null
$res = @(1..5 | ForEach-Object { $streamed.Add($_); $_ } | Tee-Object -Variable sideStore | Select-Object -First 2)
if ($res.Length -ne 2 -or $res[0] -ne 1 -or $res[1] -ne 2) {
    Write-Host "FAIL: Lazy streaming evaluation with Tee-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
