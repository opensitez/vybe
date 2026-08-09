# vybe-test: powershell/ref_parameters/ref_param_multiple_refs
function Update-Two([ref]$a, [ref]$b) {
    $a.Value += 1
    $b.Value += 2
}
$x = 10
$y = 20
Update-Two ([ref]$x) ([ref]$y)
if ($x -ne 11 -or $y -ne 22) {
    Write-Host "FAIL: multiple refs expected x=11, y=22, got x=$x, y=$y"
    exit 1
}
Write-Host "PASS"
exit 0
