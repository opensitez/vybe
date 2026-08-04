# vybe-test: powershell/arrays/create_array
$arr = @(1, 2, 3)
$count = $arr.Count
if ($count -ne 3) {
    Write-Host "FAIL: expected 3, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
