# vybe-test: powershell/collections_generic_list/capacity_expansion
$list = [System.Collections.Generic.List[int]]::new(2)
$list.Add(1); $list.Add(2); $list.Add(3)
if ($list.Capacity -lt 3 -or $list.Count -ne 3) {
    Write-Host "FAIL: Capacity auto-expansion failed"
    exit 1
}
Write-Host "PASS"
exit 0
