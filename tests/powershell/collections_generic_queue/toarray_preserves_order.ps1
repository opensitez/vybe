# vybe-test: powershell/collections_generic_queue/toarray_preserves_order
$q = [System.Collections.Generic.Queue[string]]::new()
$q.Enqueue("a"); $q.Enqueue("b"); $q.Enqueue("c")
$arr = $q.ToArray()
if ($arr.Length -ne 3 -or $arr[0] -ne "a" -or $arr[2] -ne "c") {
    Write-Host "FAIL: Queue ToArray failed"
    exit 1
}
Write-Host "PASS"
exit 0
