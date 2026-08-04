# vybe-test: powershell/loops/labeled_continue
:outer for ($i = 0; $i -lt 2; $i++) {
    $count = 0
    for ($j = 0; $j -lt 3; $j++) {
        if ($j -eq 1) {
            continue outer
        }
        $count++
    }
}
Write-Host "PASS"
exit 0
