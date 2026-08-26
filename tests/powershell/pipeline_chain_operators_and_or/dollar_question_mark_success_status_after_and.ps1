# vybe-test: powershell/pipeline_chain_operators_and_or/dollar_question_mark_success_status_after_and
function OkCmd { return $true }
OkCmd && OkCmd
if (-not $?) {
    Write-Host "FAIL: `$? should be `$true after successful && chain"
    exit 1
}
Write-Host "PASS"
exit 0
