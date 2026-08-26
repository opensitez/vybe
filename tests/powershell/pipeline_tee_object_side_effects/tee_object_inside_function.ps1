# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_inside_function
function Split-Stream {
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process {
        $Val | Tee-Object -Variable script:sideVal
    }
}
$script:sideVal = $null
$res = 42 | Split-Stream
if ($res -ne 42 -or $script:sideVal -ne 42) {
    Write-Host "FAIL: Tee-Object inside function failed"
    exit 1
}
Write-Host "PASS"
exit 0
