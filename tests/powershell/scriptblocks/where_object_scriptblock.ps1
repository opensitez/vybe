# vybe-test: powershell/scriptblocks/where_object_scriptblock
$numbers = 1..10
$evens = $numbers | Where-Object { $_ % 2 -eq 0 }
$count = ($evens | Measure-Object).Count
if ($count -ne 5) {
    Write-Host "FAIL: expected 5 evens, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
