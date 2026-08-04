# vybe-test: powershell/loops/for_loop_count
$count = 0
for ($i = 0; $i -lt 5; $i++) {
    $count++
}
if ($count -ne 5) {
    Write-Host "FAIL: expected 5, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
