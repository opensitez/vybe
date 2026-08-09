# vybe-test: powershell/generic_types/generic_read_only_collection
$list = [System.Collections.Generic.List[int]]::new()
$list.Add(5)
$ro = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
if ($ro[0] -ne 5 -or $ro.Count -ne 1) {
    Write-Host "FAIL: ReadOnlyCollection expected item 5, Count 1"
    exit 1
}
Write-Host "PASS"
exit 0
