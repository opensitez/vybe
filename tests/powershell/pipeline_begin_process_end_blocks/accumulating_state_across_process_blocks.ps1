# vybe-test: powershell/pipeline_begin_process_end_blocks/accumulating_state_across_process_blocks
function Sum-Pipeline {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Num)
    begin { $total = 0 }
    process { $total += $Num }
    end { return $total }
}
$res = 10, 20, 30, 40 | Sum-Pipeline
if ($res -ne 100) {
    Write-Host "FAIL: Accumulating state across process blocks failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
