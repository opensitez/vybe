# vybe-test: powershell/pipeline_object_capture/capture_hashtable_pipeline
$res = @{ A = 10; B = 20 }.GetEnumerator() | Where-Object { $_.Value -gt 15 }
if ($res.Key -ne "B" -or $res.Value -ne 20) {
    Write-Host "FAIL: hashtable pipeline capture expected B=20"
    exit 1
}
Write-Host "PASS"
exit 0
