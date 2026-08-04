# vybe-test: powershell/recursion/binary_search
function BinarySearch([int[]]$arr, [int]$target, [int]$lo, [int]$hi) {
    if ($lo -gt $hi) { return -1 }
    $mid = [int](($lo + $hi) / 2)
    if ($arr[$mid] -eq $target) { return $mid }
    if ($arr[$mid] -lt $target) { return BinarySearch $arr $target ($mid + 1) $hi }
    return BinarySearch $arr $target $lo ($mid - 1)
}
$sorted = @(1,3,5,7,9,11,13,15)
$idx = BinarySearch $sorted 7 0 ($sorted.Length - 1)
if ($idx -ne 3) { Write-Host "FAIL: expected index 3, got $idx"; exit 1 }
$miss = BinarySearch $sorted 6 0 ($sorted.Length - 1)
if ($miss -ne -1) { Write-Host "FAIL: missing element should return -1"; exit 1 }
Write-Host "PASS"
exit 0
