# vybe-test: powershell/ordered_hashtables/ordered_hashtable_pipeline_input
$h = [ordered]@{ P1 = 100; P2 = 200 }
$res = $h | ForEach-Object { $_.Count }
if ($res -ne 2) {
    Write-Host "FAIL: pipeline item Count expected 2, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
