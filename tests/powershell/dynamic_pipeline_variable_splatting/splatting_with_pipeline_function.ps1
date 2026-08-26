# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_with_pipeline_function
function Target-PipeSplat {
    param(
        [Parameter(ValueFromPipeline=$true)][int]$InputObject,
        [int]$Multiplier = 1
    )
    process { $InputObject * $Multiplier }
}
$opts = @{ Multiplier = 5 }
$res = @(1, 2, 3 | Target-PipeSplat @opts)
if ($res.Length -ne 3 -or $res[0] -ne 5 -or $res[2] -ne 15) {
    Write-Host "FAIL: Splatting with pipeline function failed"
    exit 1
}
Write-Host "PASS"
exit 0
