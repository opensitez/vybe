# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_multiple_hashtables
function Target-MultiSplat {
    param([string]$A, [string]$B, [string]$C, [string]$D)
    return "$A$B$C$D"
}
$p1 = @{ A = "1"; B = "2" }
$p2 = @{ C = "3"; D = "4" }
$res = Target-MultiSplat @p1 @p2
if ($res -ne "1234") {
    Write-Host "FAIL: Multiple splatted hashtables failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
