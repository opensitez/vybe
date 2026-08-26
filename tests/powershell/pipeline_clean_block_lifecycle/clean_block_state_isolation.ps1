# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_state_isolation
function Test-CleanPipeline {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$InputObject)
    begin { $total = 0 }
    process { $total += $InputObject }
    end { $total }
}
$res = 1..5 | Test-CleanPipeline
if ($res -ne 15) {
    Write-Host "FAIL: Pipeline execution failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
