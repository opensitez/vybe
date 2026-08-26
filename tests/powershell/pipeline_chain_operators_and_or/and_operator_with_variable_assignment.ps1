# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_with_variable_assignment
function Get-Data { return "Hello" }
$data = $null
(Get-Data) && ($data = "Saved")
if ($data -ne "Saved") {
    Write-Host "FAIL: && operator with variable assignment failed"
    exit 1
}
Write-Host "PASS"
exit 0
