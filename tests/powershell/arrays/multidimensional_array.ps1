# vybe-test: powershell/arrays/multidimensional_array
$arr = New-Object 'int[,]' 2, 3
$arr[0, 0] = 1
$arr[0, 1] = 2
$arr[1, 0] = 3
if ($arr[0, 0] -ne 1) {
    Write-Host "FAIL: expected arr[0,0] = 1, got $($arr[0, 0])"
    exit 1
}
if ($arr[1, 0] -ne 3) {
    Write-Host "FAIL: expected arr[1,0] = 3, got $($arr[1, 0])"
    exit 1
}
Write-Host "PASS"
exit 0
