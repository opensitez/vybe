# vybe-test: powershell/collections_generic_hashset/copyto_array
$set = [System.Collections.Generic.HashSet[int]]::new([int[]]@(10, 20, 30))
[int[]]$arr = [int[]]::new(3)
$set.CopyTo($arr)
if ($arr.Length -ne 3 -or -not ($arr -contains 10) -or -not ($arr -contains 30)) {
    Write-Host "FAIL: CopyTo array failed"
    exit 1
}
Write-Host "PASS"
exit 0
