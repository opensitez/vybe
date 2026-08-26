# vybe-test: powershell/collections_generic_list/add_elements_and_count
$list = [System.Collections.Generic.List[string]]::new()
$list.Add("apple")
$list.Add("banana")
if ($list.Count -ne 2 -or $list[0] -ne "apple" -or $list[1] -ne "banana") {
    Write-Host "FAIL: Generic List Add/Count mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
