# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_hashtable_stream
$ht = @{ a = 1; b = 2 }
$sideHt = $null
$res = $ht | Tee-Object -Variable sideHt
if ($sideHt["a"] -ne 1 -or $sideHt["b"] -ne 2) {
    Write-Host "FAIL: Tee-Object with hashtable stream failed"
    exit 1
}
Write-Host "PASS"
exit 0
