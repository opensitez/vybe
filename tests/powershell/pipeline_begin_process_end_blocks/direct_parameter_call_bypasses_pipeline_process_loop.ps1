# vybe-test: powershell/pipeline_begin_process_end_blocks/direct_parameter_call_bypasses_pipeline_process_loop
function Direct-Call {
    param([Parameter(ValueFromPipeline=$true)][int]$Num)
    begin { $b = "B" }
    process { "P:$Num" }
    end { "E" }
}
$res = @(Direct-Call -Num 42)
if ($res.Length -ne 2 -or $res[0] -ne "P:42" -or $res[1] -ne "E") {
    Write-Host "FAIL: Direct parameter call failed, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
