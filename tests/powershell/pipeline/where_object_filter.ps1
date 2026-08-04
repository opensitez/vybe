# vybe-test: powershell/pipeline/where_object_filter
$numbers = 1..10
$result = $numbers | Where-Object { $_ -gt 5 }
$count = $result.Count
if ($count -ne 5) {
    Write-Host "FAIL: expected 5, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
