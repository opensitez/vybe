# vybe-test: powershell/pipeline_object_capture/capture_in_hashtable_literal
$hash = @{
    Evens = @(1..5 | Where-Object { $_ % 2 -eq 0 })
}
if ($hash.Evens.Count -ne 2 -or $hash.Evens[1] -ne 4) {
    Write-Host "FAIL: pipeline capture in hashtable literal expected Evens=@(2,4)"
    exit 1
}
Write-Host "PASS"
exit 0
