# vybe-test: powershell/loops/do_until_loop
$count = 0
do {
    $count++
} until ($count -ge 3)
if ($count -ne 3) {
    Write-Host "FAIL: expected 3, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
