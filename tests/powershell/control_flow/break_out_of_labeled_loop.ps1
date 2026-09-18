# vybe-test: powershell/control_flow/break_out_of_labeled_loop
# A labeled break exits the specified outer loop from within an inner nested loop
$visited = @()

:outerLoop for ($i = 0; $i -lt 4; $i++) {
    for ($j = 0; $j -lt 4; $j++) {
        if ($i -eq 1 -and $j -eq 2) {
            break outerLoop
        }
        $visited += "$i-$j"
    }
}

# Should visit (0,0), (0,1), (0,2), (0,3), (1,0), (1,1) then break outerLoop completely
$expected = "0-0, 0-1, 0-2, 0-3, 1-0, 1-1"
$actual = $visited -join ", "

if ($actual -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$actual'"
    exit 1
}

Write-Host "PASS"
exit 0
