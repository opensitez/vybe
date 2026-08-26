# vybe-test: powershell/pipeline_group_object_operations/group_object_ashashtable_flag
$items = @(
    [pscustomobject]@{ Key = "K1"; Val = 10 },
    [pscustomobject]@{ Key = "K2"; Val = 20 }
)
$ht = $items | Group-Object -Property Key -AsHashTable
if ($ht["K1"].Val -ne 10 -or $ht["K2"].Val -ne 20) {
    Write-Host "FAIL: Group-Object -AsHashTable failed"
    exit 1
}
Write-Host "PASS"
exit 0
