# vybe-test: powershell/pipeline_begin_process_end_blocks/process_block_early_exit_via_break
function Stop-OnZero {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Num)
    process {
        if ($Num -eq 0) { break }
        $Num
    }
}
$res = @(1, 2, 0, 3, 4 | Stop-OnZero)
if ($res.Length -ne 2 -or $res[0] -ne 1 -or $res[1] -ne 2) {
    Write-Host "FAIL: break statement in pipeline process block failed, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
