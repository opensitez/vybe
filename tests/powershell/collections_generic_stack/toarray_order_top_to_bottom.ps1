# vybe-test: powershell/collections_generic_stack/toarray_order_top_to_bottom
$s = [System.Collections.Generic.Stack[string]]::new()
$s.Push("first"); $s.Push("second"); $s.Push("third")
$arr = $s.ToArray()
if ($arr.Length -ne 3 -or $arr[0] -ne "third" -or $arr[2] -ne "first") {
    Write-Host "FAIL: Stack ToArray top-to-bottom order failed"
    exit 1
}
Write-Host "PASS"
exit 0
