# vybe-test: powershell/loops/while_loop_count
$i = 0
$count = 0
while ($i -lt 3) {
    $count++
    $i++
}
if ($count -ne 3) {
    Write-Host "FAIL: expected 3, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
