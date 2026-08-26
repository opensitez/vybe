# vybe-test: powershell/pipeline_begin_process_end_blocks/output_from_both_process_and_end_blocks
function Append-Total {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Num)
    begin { $sum = 0 }
    process { $sum += $Num; $Num }
    end { "TOTAL:$sum" }
}
$res = @(10, 20 | Append-Total)
if ($res.Length -ne 3 -or $res[0] -ne 10 -or $res[1] -ne 20 -or $res[2] -ne "TOTAL:30") {
    Write-Host "FAIL: Output from both process and end blocks failed"
    exit 1
}
Write-Host "PASS"
exit 0
