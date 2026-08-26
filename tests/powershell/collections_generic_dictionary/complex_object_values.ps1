# vybe-test: powershell/collections_generic_dictionary/complex_object_values
$d = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[int]]]::new()
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3))
$d.Add("numbers", $list)
if ($d["numbers"].Count -ne 3 -or $d["numbers"][2] -ne 3) {
    Write-Host "FAIL: Nested List inside Dictionary failed"
    exit 1
}
Write-Host "PASS"
exit 0
