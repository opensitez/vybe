# vybe-test: powershell/arrays/array_indexing
$arr = @(10, 20, 30)
$value = $arr[1]
if ($value -ne 20) {
    Write-Host "FAIL: expected 20, got $value"
    exit 1
}
Write-Host "PASS"
exit 0
