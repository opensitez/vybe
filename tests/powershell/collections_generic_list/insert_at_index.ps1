# vybe-test: powershell/collections_generic_list/insert_at_index
$list = [System.Collections.Generic.List[string]]::new([string[]]@("a", "c"))
$list.Insert(1, "b")
if ($list.Count -ne 3 -or $list[1] -ne "b" -or $list[2] -ne "c") {
    Write-Host "FAIL: Insert failed"
    exit 1
}
Write-Host "PASS"
exit 0
