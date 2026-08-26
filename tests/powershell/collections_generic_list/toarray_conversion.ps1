# vybe-test: powershell/collections_generic_list/toarray_conversion
$list = [System.Collections.Generic.List[string]]::new([string[]]@("x", "y", "z"))
$arr = $list.ToArray()
if ($arr.GetType().IsArray -ne $true -or $arr.Length -ne 3 -or $arr[0] -ne "x") {
    Write-Host "FAIL: ToArray conversion failed"
    exit 1
}
Write-Host "PASS"
exit 0
