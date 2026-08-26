# vybe-test: powershell/pipeline_group_object_operations/group_object_ashashtable_asstring_flag
$items = @(
    [pscustomobject]@{ Code = 100; Val = "A" },
    [pscustomobject]@{ Code = 200; Val = "B" }
)
$ht = $items | Group-Object -Property Code -AsHashTable -AsString
if ($ht["100"].Val -ne "A" -or $ht["200"].Val -ne "B") {
    Write-Host "FAIL: Group-Object -AsHashTable -AsString failed"
    exit 1
}
Write-Host "PASS"
exit 0
