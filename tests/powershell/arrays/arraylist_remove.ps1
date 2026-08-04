# vybe-test: powershell/arrays/arraylist_remove
$list = [System.Collections.ArrayList]@(10, 20, 30)
$list.Remove(20)
if ($list.Count -ne 2) {
    Write-Host "FAIL: expected 2 elements, got $($list.Count)"
    exit 1
}
if ($list[1] -ne 30) {
    Write-Host "FAIL: expected element at index 1 to be 30, got $($list[1])"
    exit 1
}
Write-Host "PASS"
exit 0
