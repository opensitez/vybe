# vybe-test: powershell/cmdlets/measure_object_count
$items = "a", "b", "c"
$result = $items | Measure-Object
$count = $result.Count
if ($count -ne 3) {
    Write-Host "FAIL: expected 3, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
