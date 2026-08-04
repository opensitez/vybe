# vybe-test: powershell/operators/null_coalescing_assignment
$x = $null
$x ??= "assigned"
if ($x -ne "assigned") {
    Write-Host "FAIL: expected 'assigned', got '$x'"
    exit 1
}
$y = "existing"
$y ??= "new"
if ($y -ne "existing") {
    Write-Host "FAIL: expected 'existing', got '$y'"
    exit 1
}
Write-Host "PASS"
exit 0
