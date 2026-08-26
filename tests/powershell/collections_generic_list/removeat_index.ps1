# vybe-test: powershell/collections_generic_list/removeat_index
$list = [System.Collections.Generic.List[string]]::new([string[]]@("first", "middle", "last"))
$list.RemoveAt(1)
if ($list.Count -ne 2 -or $list[1] -ne "last") {
    Write-Host "FAIL: RemoveAt failed"
    exit 1
}
Write-Host "PASS"
exit 0
