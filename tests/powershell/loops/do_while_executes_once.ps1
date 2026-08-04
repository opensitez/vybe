# vybe-test: powershell/loops/do_while_executes_once
$count = 0
do {
    $count++
} while ($false)
if ($count -ne 1) {
    Write-Host "FAIL: expected 1, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
