# vybe-test: powershell/parameters_validate_range/validaterange_with_pipeline_input
function Test-RangePipe {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [ValidateRange(10, 20)]
        [int]$Val
    )
    process { $Val * 2 }
}
$res = 15 | Test-RangePipe
if ($res -ne 30) {
    Write-Host "FAIL: ValidateRange pipeline input failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
