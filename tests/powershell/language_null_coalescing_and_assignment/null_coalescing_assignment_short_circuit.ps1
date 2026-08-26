# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_assignment_short_circuit
$evalCount = 0
function Get-ExpensiveDefault {
    $script:evalCount++
    return "Expensive"
}
$val = "AlreadySet"
$val ??= (Get-ExpensiveDefault)
if ($val -ne "AlreadySet" -or $evalCount -ne 0) {
    Write-Host "FAIL: ??= should not evaluate RHS function when target is set"
    exit 1
}
Write-Host "PASS"
exit 0
