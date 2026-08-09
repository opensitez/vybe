# vybe-test: powershell/generic_types/generic_list_instantiation
$list = [System.Collections.Generic.List[int]]::new()
$list.Add(10)
$list.Add(20)
if ($list.Count -ne 2 -or $list[0] -ne 10 -or $list[1] -ne 20) {
    Write-Host "FAIL: Generic List[int] expected 10, 20"
    exit 1
}
Write-Host "PASS"
exit 0
