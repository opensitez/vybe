# vybe-test: powershell/pipeline_begin_process_end_blocks/process_block_skip_via_continue
function Skip-Negatives {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Num)
    process {
        if ($Num -lt 0) { return } # return in process block skips current item
        $Num
    }
}
$res = @(1, -5, 2, -3, 3 | Skip-Negatives)
if ($res.Length -ne 3 -or $res[0] -ne 1 -or $res[1] -ne 2 -or $res[2] -ne 3) {
    Write-Host "FAIL: Skip via return in process block failed, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
