# vybe-test: powershell/exceptions_trap_statement_scope/trap_in_pipeline_process_block
$trapCount = 0
function Test-PipeTrap {
    param([Parameter(ValueFromPipeline=$true)][int]$N)
    process {
        trap {
            $script:trapCount++
            continue
        }
        if ($N -eq 2) { 1 / 0 }
        $N
    }
}
$res = @(1, 2, 3 | Test-PipeTrap)
if ($trapCount -ne 1 -or $res.Length -lt 2) {
    Write-Host "FAIL: Trap in pipeline process block failed, count=$trapCount"
    exit 1
}
Write-Host "PASS"
exit 0
