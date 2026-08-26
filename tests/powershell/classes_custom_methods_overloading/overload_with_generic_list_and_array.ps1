# vybe-test: powershell/classes_custom_methods_overloading/overload_with_generic_list_and_array
class ListArrayOverloadTarget {
    [string]Process([System.Collections.Generic.List[int]]$l) { return "List" }
    [string]Process([int[]]$a) { return "Array" }
}
$t = [ListArrayOverloadTarget]::new()
$list = [System.Collections.Generic.List[int]]::new()
[int[]]$arr = [int[]]@(1, 2)
if ($t.Process($list) -ne "List" -or $t.Process($arr) -ne "Array") {
    Write-Host "FAIL: Generic List vs Array overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
