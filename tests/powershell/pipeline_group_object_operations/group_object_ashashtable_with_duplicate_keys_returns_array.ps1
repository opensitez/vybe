# vybe-test: powershell/pipeline_group_object_operations/group_object_ashashtable_with_duplicate_keys_returns_array
$items = @(
    [pscustomobject]@{ GroupKey = "G1"; Val = 10 },
    [pscustomobject]@{ GroupKey = "G1"; Val = 20 }
)
$ht = $items | Group-Object -Property GroupKey -AsHashTable
if ($ht["G1"].Count -ne 2 -or $ht["G1"][0].Val -ne 10 -or $ht["G1"][1].Val -ne 20) {
    Write-Host "FAIL: Group-Object -AsHashTable with multi-item group failed"
    exit 1
}
Write-Host "PASS"
exit 0
