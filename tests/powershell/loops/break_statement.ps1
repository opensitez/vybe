# vybe-test: powershell/loops/break_statement
$count = 0
for ($i = 0; $i -lt 10; $i++) {
    if ($i -eq 3) {
        break
    }
    $count++
}
if ($count -ne 3) {
    Write-Host "FAIL: expected 3, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
