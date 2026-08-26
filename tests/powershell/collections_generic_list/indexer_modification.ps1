# vybe-test: powershell/collections_generic_list/indexer_modification
$list = [System.Collections.Generic.List[string]]::new([string[]]@("old", "keep"))
$list[0] = "new"
if ($list[0] -ne "new" -or $list[1] -ne "keep") {
    Write-Host "FAIL: Indexer assignment failed"
    exit 1
}
Write-Host "PASS"
exit 0
