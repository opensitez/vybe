# vybe-test: powershell/parameters_validate_count/validatecount_with_pipeline_input_array
function Collect-Pipe {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [ValidateCount(2, 4)]
        [string[]]$Tags
    )
    process { $Tags.Length }
}
$res = ,@("tag1", "tag2") | Collect-Pipe
if ($res -ne 2) {
    Write-Host "FAIL: ValidateCount pipeline input failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
