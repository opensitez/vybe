# vybe-test: powershell/pipeline_begin_process_end_blocks/execution_order_begin_process_end
function Test-BlockOrder {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$InputObject)
    begin { $events = [System.Collections.Generic.List[string]]::new(); $events.Add("BEGIN") }
    process { $events.Add("PROC:$InputObject") }
    end { $events.Add("END"); return ($events -join "->") }
}
$res = 1, 2, 3 | Test-BlockOrder
if ($res -ne "BEGIN->PROC:1->PROC:2->PROC:3->END") {
    Write-Host "FAIL: Block execution order failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
