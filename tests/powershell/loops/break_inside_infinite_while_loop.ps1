# vybe-test: powershell/loops/break_inside_infinite_while_loop
# A while ($true) loop is safely exited upon encountering an explicit break condition
$iteration = 0

while ($true) {
    $iteration++
    if ($iteration -eq 7) {
        break
    }
}

if ($iteration -ne 7) {
    Write-Host "FAIL: expected loop to break at iteration 7, got $iteration"
    exit 1
}

Write-Host "PASS"
exit 0
