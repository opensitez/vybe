# vybe-test: powershell/pipeline_begin_process_end_blocks/begin_block_runs_once_even_with_no_input
function Test-EmptyInput {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$InputObject)
    begin { $b = 1 }
    process { $p = 1 }
    end { return "B:$b,P:$p" }
}
$res = @() | Test-EmptyInput
if ($res -ne "B:1,P:") {
    Write-Host "FAIL: Empty pipeline input begin/process/end failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
