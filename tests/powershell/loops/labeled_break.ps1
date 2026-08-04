# vybe-test: powershell/loops/labeled_break
:outer for ($i = 0; $i -lt 3; $i++) {
    for ($j = 0; $j -lt 3; $j++) {
        if ($j -eq 1) {
            break outer
        }
    }
}
if ($i -ne 0) {
    Write-Host "FAIL: expected i = 0, got $i"
    exit 1
}
Write-Host "PASS"
exit 0
