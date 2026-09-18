# vybe-test: powershell/control_flow/continue_in_labeled_loop
# A labeled continue skips remainder of inner loop and triggers the next iteration of specified outer loop
$visited = @()

:mainLoop for ($row = 0; $row -lt 3; $row++) {
    for ($col = 0; $col -lt 3; $col++) {
        if ($col -eq 1) {
            continue mainLoop
        }
        $visited += "${row}-${col}"
    }
}

# In each row, col=0 is recorded, then col=1 immediately jumps to next row iteration
$expected = "0-0, 1-0, 2-0"
$actual = $visited -join ", "

if ($actual -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$actual'"
    exit 1
}

Write-Host "PASS"
exit 0
