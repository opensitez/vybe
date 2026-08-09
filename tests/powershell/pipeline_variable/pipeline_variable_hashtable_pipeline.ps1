# vybe-test: powershell/pipeline_variable/pipeline_variable_hashtable_pipeline
$hashes = @(@{ K = "A" }, @{ K = "B" })
$res = $hashes | ForEach-Object -PipelineVariable h { $h["K"] } | ForEach-Object { "$($h.K):$_" }
if ($res[0] -ne "A:A" -or $res[1] -ne "B:B") {
    Write-Host "FAIL: hashtable PipelineVariable binding expected A:A, B:B"
    exit 1
}
Write-Host "PASS"
exit 0
