# vybe-test: powershell/collections_generic_list/clear_elements
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3))
$list.Clear()
if ($list.Count -ne 0) {
    Write-Host "FAIL: Clear failed"
    exit 1
}
Write-Host "PASS"
exit 0
