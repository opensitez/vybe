# vybe-test: powershell/language_ternary_conditional_operator/ternary_false_branch_evaluation
$condition = $false
$result = $condition ? "TrueBranch" : "FalseBranch"
if ($result -ne "FalseBranch") {
    Write-Host "FAIL: Ternary false branch failed, got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
