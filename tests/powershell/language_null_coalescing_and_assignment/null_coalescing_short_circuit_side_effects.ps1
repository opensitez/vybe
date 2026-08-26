# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_short_circuit_side_effects
$sideEffect = $false
$left = "Valid"
$res = $left ?? ($sideEffect = $true)
if ($res -ne "Valid" -or $sideEffect -ne $false) {
    Write-Host "FAIL: Null coalescing RHS should not evaluate when LHS is not null"
    exit 1
}
Write-Host "PASS"
exit 0
