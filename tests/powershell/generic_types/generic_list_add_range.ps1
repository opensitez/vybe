# vybe-test: powershell/generic_types/generic_list_add_range
$list = [System.Collections.Generic.List[string]]::new()
$list.AddRange([string[]]@("A", "B", "C"))
if ($list.Count -ne 3 -or $list[2] -ne "C") {
    Write-Host "FAIL: AddRange expected Count=3, last item 'C'"
    exit 1
}
Write-Host "PASS"
exit 0
