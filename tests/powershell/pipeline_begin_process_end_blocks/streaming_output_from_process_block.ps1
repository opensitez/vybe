# vybe-test: powershell/pipeline_begin_process_end_blocks/streaming_output_from_process_block
function Double-Stream {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Num)
    process { $Num * 2 }
}
$res = @(1, 2, 3 | Double-Stream)
if ($res.Length -ne 3 -or $res[0] -ne 2 -or $res[2] -ne 6) {
    Write-Host "FAIL: Streaming output from process block failed"
    exit 1
}
Write-Host "PASS"
exit 0
