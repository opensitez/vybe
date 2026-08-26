# vybe-test: powershell/language_ternary_conditional_operator/ternary_short_circuit_side_effects
$sideEffect = $false
$result = $true ? "OK" : ($sideEffect = $true)
if ($result -ne "OK" -or $sideEffect -ne $false) {
    Write-Host "FAIL: Ternary false branch should not evaluate when condition is true"
    exit 1
}
Write-Host "PASS"
exit 0
