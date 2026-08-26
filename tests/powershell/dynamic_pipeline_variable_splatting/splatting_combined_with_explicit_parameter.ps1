# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_combined_with_explicit_parameter
function Target-MixFunc {
    param([string]$A, [string]$B, [string]$C)
    return "$A-$B-$C"
}
$p = @{ B = "2"; C = "3" }
$res = Target-MixFunc -A "1" @p
if ($res -ne "1-2-3") {
    Write-Host "FAIL: Combining splatting with explicit parameter failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
