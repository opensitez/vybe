# vybe-test: powershell/generic_types/generic_nested_list_of_lists
$matrix = [System.Collections.Generic.List[System.Collections.Generic.List[int]]]::new()
$row = [System.Collections.Generic.List[int]]::new()
$row.Add(99)
$matrix.Add($row)
if ($matrix[0][0] -ne 99) {
    Write-Host "FAIL: nested List[List[int]] matrix[0][0] expected 99, got $($matrix[0][0])"
    exit 1
}
Write-Host "PASS"
exit 0
