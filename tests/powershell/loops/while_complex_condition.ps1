# vybe-test: powershell/loops/while_complex_condition
$x = 1
$y = 100
while ($x -lt 50 -and $y -gt 60) {
    $x *= 2
    $y -= 10
}
# x: 1→2→4→8→16→32→64 (stops when x=64 >= 50)
# y: 100→90→80→70→60 (stops when y=60, not > 60)
# After 5 iterations: x=32, y=60; next check: 32<50 AND 60>60 → false (60 not > 60), loop exits
if ($x -ne 64) { Write-Host "FAIL: x=$x expected 64"; exit 1 }
if ($y -ne 60) { Write-Host "FAIL: y=$y expected 60"; exit 1 }
Write-Host "PASS"
exit 0
