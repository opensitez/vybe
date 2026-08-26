# vybe-test: powershell/language_ternary_conditional_operator/ternary_true_branch_evaluation
$condition = $true
$result = $condition ? "TrueBranch" : "FalseBranch"
if ($result -ne "TrueBranch") {
    Write-Host "FAIL: Ternary true branch failed, got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
