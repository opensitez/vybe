# vybe-test: powershell/collections_generic_list/indexof_element
$list = [System.Collections.Generic.List[int]]::new([int[]]@(100, 200, 300))
if ($list.IndexOf(200) -ne 1 -or $list.IndexOf(999) -ne -1) {
    Write-Host "FAIL: IndexOf failed"
    exit 1
}
Write-Host "PASS"
exit 0
