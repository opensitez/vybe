# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_inside_if_condition
function Check1 { return $true }
function Check2 { return $true }
$ran = $false
if ((Check1) && (Check2)) {
    $ran = $true
}
if (-not $ran) {
    Write-Host "FAIL: && operator inside if condition failed"
    exit 1
}
Write-Host "PASS"
exit 0
