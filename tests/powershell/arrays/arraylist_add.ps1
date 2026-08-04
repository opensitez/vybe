# vybe-test: powershell/arrays/arraylist_add
$list = [System.Collections.ArrayList]@()
[void]$list.Add(10)
[void]$list.Add(20)
[void]$list.Add(30)
if ($list.Count -ne 3) {
    Write-Host "FAIL: expected 3 elements, got $($list.Count)"
    exit 1
}
if ($list[1] -ne 20) {
    Write-Host "FAIL: expected element at index 1 to be 20, got $($list[1])"
    exit 1
}
Write-Host "PASS"
exit 0
