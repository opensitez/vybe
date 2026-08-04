# vybe-test: powershell/variables/multiple_assignment_swap
$a = 1
$b = 2
$a, $b = $b, $a
if ($a -ne 2) { Write-Host "FAIL: a=$a expected 2"; exit 1 }
if ($b -ne 1) { Write-Host "FAIL: b=$b expected 1"; exit 1 }
# Destructure array
$x, $y, $z = @(10, 20, 30)
if ($x -ne 10) { Write-Host "FAIL: x"; exit 1 }
if ($y -ne 20) { Write-Host "FAIL: y"; exit 1 }
if ($z -ne 30) { Write-Host "FAIL: z"; exit 1 }
Write-Host "PASS"
exit 0
