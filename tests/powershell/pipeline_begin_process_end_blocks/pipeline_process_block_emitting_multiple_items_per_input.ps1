# vybe-test: powershell/pipeline_begin_process_end_blocks/pipeline_process_block_emitting_multiple_items_per_input
function Expand-Item {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][string]$Name)
    process {
        "$Name-1"
        "$Name-2"
    }
}
$res = @("A", "B" | Expand-Item)
if ($res.Length -ne 4 -or $res[0] -ne "A-1" -or $res[3] -ne "B-2") {
    Write-Host "FAIL: Multi-item emission per process block failed"
    exit 1
}
Write-Host "PASS"
exit 0
