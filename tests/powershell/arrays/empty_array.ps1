# vybe-test: powershell/arrays/empty_array
$arr = @()
$count = $arr.Count
if ($count -ne 0) {
    Write-Host "FAIL: expected 0, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
