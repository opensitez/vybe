# vybe-test: powershell/arrays/array_clone
$arr = @(1, 2, 3)
$clone = $arr.Clone()
$clone[0] = 99
if ($arr[0] -ne 1) {
    Write-Host "FAIL: original array should not be modified"
    exit 1
}
if ($clone[0] -ne 99) {
    Write-Host "FAIL: cloned array should have new value"
    exit 1
}
Write-Host "PASS"
exit 0
