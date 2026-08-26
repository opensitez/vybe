# vybe-test: powershell/pipeline_begin_process_end_blocks/multiple_functions_chained_in_pipeline
function Step1 {
    param([Parameter(ValueFromPipeline=$true)][int]$N)
    process { $N + 1 }
}
function Step2 {
    param([Parameter(ValueFromPipeline=$true)][int]$N)
    process { $N * 3 }
}
$res = @(1, 2, 3 | Step1 | Step2) # (1+1)*3=6, (2+1)*3=9, (3+1)*3=12
if ($res.Length -ne 3 -or $res[0] -ne 6 -or $res[1] -ne 9 -or $res[2] -ne 12) {
    Write-Host "FAIL: Chained pipeline functions failed"
    exit 1
}
Write-Host "PASS"
exit 0
