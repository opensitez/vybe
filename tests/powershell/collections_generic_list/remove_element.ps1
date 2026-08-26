# vybe-test: powershell/collections_generic_list/remove_element
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3, 2))
$removed = $list.Remove(2)
if (-not $removed -or $list.Count -ne 3 -or $list[1] -ne 3) {
    Write-Host "FAIL: Remove first matching element failed"
    exit 1
}
Write-Host "PASS"
exit 0
